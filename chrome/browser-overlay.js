/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <http://www.gnu.org/licenses/>.
    
    ***** END LICENSE BLOCK *****
*/
const ZOTEROWINWORDINTEGRATION_ID = "zoteroWinWordIntegration@zotero.org";
const ZOTEROWINWORDINTEGRATION_PREF = "extensions.zoteroWinWordIntegration.version";

function zoteroWinWordIntegration_firstRun() {
	try {
		// find Word Startup folder (see http://support.microsoft.com/kb/210860)
		
		// first check the registry for a custom startup folder
		var path = null;
		var wrk = Components.classes["@mozilla.org/windows-registry-key;1"]
			.createInstance(Components.interfaces.nsIWindowsRegKey);
		for(var i=9; i<=13; i++) {
			try {
				wrk.open(wrk.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\9.0\\Word\\Options",
					wrk.ACCESS_READ);
				try {
					var path = wrk.readStringValue("STARTUP-PATH");
				} finally {
					wrk.close();
				}
			} catch(e) {}
			if(path) break;
		}
		
		if(path) {
			// create nsIFile from path in registry
			var startupFolder = Components.classes["@mozilla.org/file/local;1"].
				createInstance(Components.interfaces.nsILocalFile);
			startupFolder.initWithPath(path);
		} else {
			// if not in the registry, append Microsoft/Word/Startup to %AppData% (default location)
			var startupFolder = Components.classes["@mozilla.org/file/directory_service;1"]
				 .getService(Components.interfaces.nsIProperties)
				 .get("AppData", Components.interfaces.nsILocalFile);
			startupFolder.appendRelativePath("Microsoft\\Word\\Startup");
		}
		
		var oldDot = startupFolder.clone().QueryInterface(Components.interfaces.nsILocalFile);
		oldDot.append("Zotero.dot");
		if(oldDot.exists()) oldDot.remove(false);
		
		// get Zotero.dot file
		var dot = Components.classes["@mozilla.org/extensions/manager;1"].
			getService(Components.interfaces.nsIExtensionManager).
			getInstallLocation(ZOTEROWINWORDINTEGRATION_ID).
			getItemLocation(ZOTEROWINWORDINTEGRATION_ID);
		dot.append("install");
		dot.append("Zotero.dot");
		
		// copy Zotero.dot file to Word Startup folder
		dot.copyTo(startupFolder, "Zotero.dot");
	} catch(e) {
		var prompts = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
			.getService(Components.interfaces.nsIPromptService)
			.alert(null, 'Zotero Word Integration Error',
			'Zotero Word Integration could not complete installation because an error occurred. Please ensure that Word is closed, then restart Firefox.');
		throw e;
	}
}

var ext = Components.classes['@mozilla.org/extensions/manager;1']
   .getService(Components.interfaces.nsIExtensionManager).getItemForID(ZOTEROWINWORDINTEGRATION_ID);
var zoteroWinWordIntegration_prefService = Components.classes["@mozilla.org/preferences-service;1"].
	getService(Components.interfaces.nsIPrefBranch);
if(zoteroWinWordIntegration_prefService.getCharPref(ZOTEROWINWORDINTEGRATION_PREF) != ext.version) {
	zoteroWinWordIntegration_firstRun();
	zoteroWinWordIntegration_prefService.setCharPref(ZOTEROWINWORDINTEGRATION_PREF, ext.version);
}