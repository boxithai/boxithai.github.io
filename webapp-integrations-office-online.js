/**
 * @fileoverview Module for Office Online integration
 * @author kszubzda
 */

/*global Box*/

Box.Application.addModule('office-online', function(context) {

	'use strict';

	//--------------------------------------------------------------------------
	// Constants
	//--------------------------------------------------------------------------
	var APP_LOADING_STATUS_MESSAGE_ID = 'App_LoadingStatus';
	var POST_MESSAGE_READY_MESSAGE_ID = 'Host_PostmessageReady';
	var FILE_RENAME_MESSAGE_ID = 'File_Rename';
	var IFRAME_LOAD_TIME_EVENT = 'IFRAME_LOAD_TIME';
	var EDITOR_LOAD_TIME_EVENT = 'EDITOR_LOAD_TIME';

	var FRAME_HOLDER_CLASS = 'office-online-frame-holder';
	var EDITOR_IFRAME_NAME = 'office-editor';
	var EDITOR_IFRAME_CLASS = 'office-editor-frame';
	var OFFICE_ONLINE_CATEGORY_NAME = 'office_online';
	var DISABLE_SAME_ORIGIN_KEY = 'disableAllowSameOrigin';


	//--------------------------------------------------------------------------
	// Private
	//--------------------------------------------------------------------------

	var domService,
		logger,
		performance,
		startTs,
		win;

	/**
	 * Logs the load time for the Office Online iframe or editor. Note: this is measuring
	 * the load time of a Microsoft-owned iframe and therefore we cannot use
	 * boomerang.
	 *
	 * @param {string} eventName
	 * @returns {void}
	 */
	function logLoadTime(eventName) {
		var loadTime = performance.now() - startTs;
		var eventType = 'perf';
		var serviceId = context.getConfig('serviceId');
		var officeOnlineAppType = context.getConfig('officeOnlineAppType');

		/* eslint-disable camelcase */
		var params = {
			event_name: eventName,
			load_time: loadTime,
			service_id: serviceId,
			office_online_app_type: officeOnlineAppType
		};
		/* eslint-enable camelcase */
		logger.sendLog(OFFICE_ONLINE_CATEGORY_NAME, eventType, params);

		if (loadTime >= 10000) {
			// send a separate event for slow loads (> 10 seconds)
			params.event_name = params.event_name + '_slow';
			logger.sendLog(OFFICE_ONLINE_CATEGORY_NAME, eventType, params);
		}
	}

	/**
	 * Checks if name is a query parameter in the given url
	 *
	 * @param {string} name The name of the query parameter to search for
	 * @param {string} url The URL to search in
	 * @returns {string|null}
	 */
	function getParameterByName(name, url) {
	    name = name.replace(/[\[\]]/g, '\\$&');
	    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
	        results = regex.exec(url);
	    if (!results) return null;
	    if (!results[2]) return '';
	    return decodeURIComponent(results[2].replace(/\+/g, ' '));
	}

	//--------------------------------------------------------------------------
	// Public
	//--------------------------------------------------------------------------

	var officeOnlineModule = {
		/**
		 * Initializes the module.
		 *
		 * @returns {void}
		 */
		init: function() {
			domService = context.getService('dom');
			logger = context.getService('logger');

			// global vars for adding/removing event listeners
			win = context.getGlobal('window');

			performance = context.getGlobal('performance');
			startTs = performance.now();

			this.drawIframe();
			this.postTokenToIframe();
			this.subscribeToEventsFromOfficeOnline();
		},

		/**
		 * Destroys the module.
		 *
		 * @returns {void}
		 */
		destroy: function() {
			domService.off(win, 'message', this.receiveEventFromOfficeOnline);
		},

		/**
		 * Draws the iframe.
		 *
		 * @returns {void}
		 */
		drawIframe: function() {
			var officeOnlineFrameHolder = domService.query('.' + FRAME_HOLDER_CLASS);
			var officeEditorFrame = this.getOfficeEditorFrameElement();
			officeOnlineFrameHolder.appendChild(officeEditorFrame);
			logLoadTime(IFRAME_LOAD_TIME_EVENT);
		},

		/**
		 * Helper function for drawIframe. Returns the HTMLElement of the iframe.
		 *
		 * @returns {HTMLIFrameElement} the iframe element
		 */
		getOfficeEditorFrameElement: function() {
			var officeEditorFrame = document.createElement('iframe');
			officeEditorFrame.name = EDITOR_IFRAME_NAME;
			officeEditorFrame.className = EDITOR_IFRAME_CLASS;
			// allows true fullscreen mode in slideshow view for PowerPoint Online
			officeEditorFrame.setAttribute('allowfullscreen', 'true');
            officeEditorFrame.setAttribute('allow', 'microphone');
            officeEditorFrame.setAttribute('src', invocationUrl);
			var disableAllowSameOrigin = getParameterByName(DISABLE_SAME_ORIGIN_KEY, window.location.href);
			if (disableAllowSameOrigin) {
				officeEditorFrame.setAttribute('sandbox', 'allow-scripts allow-forms allow-popups allow-top-navigation allow-popups-to-escape-sandbox allow-downloads');
			} else {
				officeEditorFrame.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-popups allow-top-navigation allow-popups-to-escape-sandbox allow-downloads');
			}
			return officeEditorFrame;
		},

		/**
		 * Performs a POST to the iframe to invoke the Office Online editor with an access token.
		 *
		 * @returns {void}
		 */
		postTokenToIframe: function() {
			var officeForm = domService.query('.office-form');
			officeForm.submit();
		},

		/**
		 * Subscribes to events coming from the popup window.
		 *
		 * @returns {void}
		 */
		subscribeToEventsFromOfficeOnline: function() {
			domService.on(win, 'message', this.receiveEventFromOfficeOnline);
		},

		/**
		 * Decides how to handle each message received from Office Online
		 *
		 * @param {Event} event - the event from the popup
		 * @returns {void}
		 */
		receiveEventFromOfficeOnline: function(event) {
			if (event.originalEvent.data) {
				var data = JSON.parse(event.originalEvent.data);
				if (data && data.MessageId === APP_LOADING_STATUS_MESSAGE_ID) {
					logLoadTime(EDITOR_LOAD_TIME_EVENT);
					officeOnlineModule.publishReadyMessage();
				}

				if (data && data.MessageId === FILE_RENAME_MESSAGE_ID) {
					var msgValue = data.Values;
					var fileExtension = context.getConfig('fileExtension');
					var titleWithNoFilename = context.getConfig('titleWithNoFilename');
					var newName = msgValue.NewName + '.' + fileExtension;

					// @L10N this is displayed as the title of the Office Online editing window.
					// @L10N %1 is the file name
					// @L10N %2 is the translated string 'Box for Office Online - powered by Box'
					var newTitle = $t('%1 on %2', 'office_online_window_title_with_file_name', newName, titleWithNoFilename);
					window.parent.document.title = newTitle;
				}
			}
		},

		/**
		 * Gets the current timestamp
		 *
		 * @returns {number} the timestamp
		 */
		getCurrentTimestamp: function() {
			return (new Date()).getTime();
		},

		/**
		 * Publishes a message to the Office Online editor frame
		 *
		 * This messaging will happen across domains
		 *
		 * @param {Object} message - the message to pass
		 * @returns {void}
		 */
		publishMessage: function(message) {
			message = JSON.stringify(message); // IE can only accept strings in postMessage
			var officeFrame = domService.query('.office-editor-frame');
			var origin = context.getConfig('origin');
			officeFrame.contentWindow.postMessage(message, origin);
		},

		/**
		 * Sends a ready message to Office Online
		 *
		 * @returns {void}
		 */
		publishReadyMessage: function() {
			this.publishMessage({
				MessageId: POST_MESSAGE_READY_MESSAGE_ID,
				SendTime: this.getCurrentTimestamp(),
				Values: {}
			});
		}
	};

	return officeOnlineModule;
});

