
/**
 * Microsoft Office Communicator Presence Plugin.
 *
 * This jQuery plugin adds Microsoft Office Communicator presence awareness to elements.
 *
 * @option     msocpresence.displayOOUI       Whether or not to display the MSOC OOUI widget
 *
 *
 * @note
 * 		This plugin utilizes the "Name.NameCtrl" ActiveX control.  It requires that Microsoft
 *      Office Communicator be installed, can only run in Internet Explorer, and only shows
 *      presence information when served from a "trusted" site.  The main audience for a
 *      plugin like this are those developing for an intranet that uses the MS Office suite
 *      of collaboration tools. 
 *
 *      If the user's environment does NOT meet the aforementioned requirements he/she will
 *      simply not see presence information.  This plugin should not interfere with other
 *      functionality within the page.
 *
 * @usage
 *      This plugin will automatically apply presence information once the document has been
 *      fully loaded.
 *
 *      NOTE:  
 *        The OOUI widget can be disabled by setting "$.msocpresence.displayOOUI" to false
 *
 * The following CSS classes will also need to be defined:
 *
 *  - status-online
 *  - status-offline
 *  - status-away
 *  - status-busy
 *  - status-brb
 *  - status-phone
 *  - status-lunch
 *
 *
 * @example HTML Markup
 *
 *      <ul>
 *        <li class="msocuser" data-msocusername="dburdick@example.com">David Burdick</li>
 *        <li class="msocuser" data-msocusername="user1@example.com">User One</li>
 *        <li class="msocuser" data-msocusername="user2@example.com">User Two</li>
 *      </ul>
 *
 *      <a class="msocuser" data-msocusername="dburdick@example.com" href="http://david.example.com">David</a>
 *
 *
 * Copyright (c) 2011 David Burdick
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 *
 * @author				David Burdick
 *
 * @version             1.0.2
 *
 * @changelog
 *      + 1.0.0         First release
 *      + 1.0.1         ActiveXObject was breaking compatibility with non-IE 
 *                        browsers.  Fixed.
 *      + 1.0.2         Modified to automatically apply presence information
 *                        upon document load.
 *
 */

var MSOCConstants = 
{
	// Class name for any element containing an MSOC user
	MSOC_USER_CLASS: "msocuser",
	
	// Attribute containing MSOC user name (e.g. user@example.com)
	MSOC_DATA_ATTR: "data-msocusername",
	
	// The statuses that are sent to the "onStatusChange" event handler
	ONLINE_STATUS: 0,
	OFFLINE_STATUS: 1,
	AWAY_STATUS: 2,
	BUSY_STATUS: 3,
	BRB_STATUS: 4,
	PHONE_STATUS: 5,
	LUNCH_STATUS: 6,	
	
	// The classes that represent each status
	STATUS_CLASS: ["status-online",
					"status-offline",
					"status-away",
					"status-busy",
					"status-brb",
					"status-phone",
					"status-lunch"]
};

var MSOCStatusUtil = {
	
	nameCtrl: window.ActiveXObject ? new ActiveXObject('Name.NameCtrl.1') : null,
	
	/**
	 * Removes all status classes from the given element
	 *
	 * @param element the element to modify
	 */
	removeAllStatusClasses: function(element)
	{
		$(element).removeClass(MSOCConstants.STATUS_CLASS.join(" "));
	},
	
	/**
	 * Called when a user's status changes
	 * 
	 * @param {String} name The MSOC username
	 * @param {int} status The updated status
	 * @param {String} id For the purposes of this plugin, we ignore this value. 
	 */
	onStatusChange: function(name, status, id)
	{
		// Retrieve only the elements that pertain to "name"
		var elements = $("." + MSOCConstants.MSOC_USER_CLASS).filter('[' + MSOCConstants.MSOC_DATA_ATTR + '=\"' + name + '\"]');
		
		elements.each(function() 
		{
			// Clear out any old statuses
			MSOCStatusUtil.removeAllStatusClasses(this);
			
			// Add the appropriate class for the new status			
			$(this).addClass(MSOCConstants.STATUS_CLASS[status]);
		});
	}
};

(function( $ )
{
	$.msocpresence = { displayOOUI: true };
	
	var addMSOCPresence = function(options) 
	{	
		var settings = {
			// Whether or not to display the OOUI widget
			displayOOUI: false
		};
		
		if(options)
		{
			$.extend(settings, options);
		}
	
		if(MSOCStatusUtil.nameCtrl && MSOCStatusUtil.nameCtrl.PresenceEnabled)
		{
			MSOCStatusUtil.nameCtrl.OnStatusChange = MSOCStatusUtil.onStatusChange;
			
			$("." + MSOCConstants.MSOC_USER_CLASS).each(function()
			{
				// Register to receive updates for every "msocuser"
				MSOCStatusUtil.nameCtrl.GetStatus($(this).attr(MSOCConstants.MSOC_DATA_ATTR), "1");
								
				if(settings.displayOOUI)
				{
					$(this).mouseover(function() 
					{
						// Not exactly elegant...but 19 is about the size of
						// the OOUI widget.  This keeps it from overlapping
						// the element.
						var offsetX = $(this).offset().left - 19;
						
						// The OOUI position is always relative to the window,
						// regardless of where the element is.  With this in
						// mind, we need to account for any scrolling that's
						// happened.
						var offsetY = $(this).offset().top - $(window).scrollTop();
						
						MSOCStatusUtil.nameCtrl.ShowOOUI($(this).attr(MSOCConstants.MSOC_DATA_ATTR), 1, offsetX, offsetY);
					});
					
					$(this).mouseout(function() 
					{
						MSOCStatusUtil.nameCtrl.HideOOUI();
					});
				}
			});
		}
	};
	
	$(document).ready(function()
	{
		addMSOCPresence({displayOOUI: $.msocpresence.displayOOUI});
	});
	
})( jQuery );

