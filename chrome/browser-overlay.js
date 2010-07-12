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

function ZoteroWinWordIntegration_checkVersion(name, url, id, minVersion) {
	// check Zotero version
	try {
		var ext = Components.classes['@mozilla.org/extensions/manager;1']
		   .getService(Components.interfaces.nsIExtensionManager).getItemForID(id);
		var comp = Components.classes["@mozilla.org/xpcom/version-comparator;1"]
			.getService(Components.interfaces.nsIVersionComparator)
			.compare(ext.version, minVersion);
	} catch(e) {
		var comp = -1;
	}
	
	if(comp < 0) {
		var err = 'This version of Zotero WinWord Integration requires '+name+' '+minVersion+
			' or later to run. Please download the latest version of '+name+' from '+url+'.';
		var prompts = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
			.getService(Components.interfaces.nsIPromptService)
			.alert(null, 'Zotero WinWord Integration Error', err);
		throw err;
	}
}

function ZoteroWinWordIntegration_firstRun() {
	ZoteroWinWordIntegration_checkVersion("Zotero", "zotero.org", "zotero@chnm.gmu.edu", "2.1a1.SVN");
	
	try {		
		// get Zotero.dot file
		var dot = Components.classes["@mozilla.org/extensions/manager;1"].
			getService(Components.interfaces.nsIExtensionManager).
			getInstallLocation(ZOTEROWINWORDINTEGRATION_ID).
			getItemLocation(ZOTEROWINWORDINTEGRATION_ID);
		dot.append("install");
		dot.append("Zotero.dot");

		// find Word Startup folders (see http://support.microsoft.com/kb/210860)
		var appData = Components.classes["@mozilla.org/file/directory_service;1"]
				.getService(Components.interfaces.nsIProperties)
				.get("AppData", Components.interfaces.nsILocalFile);
		
		// first check the registry for a custom startup folder
		var startupFolders = [];
		var addDefaultStartupFolder = false;
		var wrk = Components.classes["@mozilla.org/windows-registry-key;1"]
			.createInstance(Components.interfaces.nsIWindowsRegKey);
		for(var i=9; i<=13; i++) {
			var path = null;
			try {
				wrk.open(Components.interfaces.nsIWindowsRegKey.ROOT_KEY_CURRENT_USER,
					"Software\\Microsoft\\Office\\"+i+".0\\Word\\Options",
					Components.interfaces.nsIWindowsRegKey.ACCESS_READ);
				try {
					path = wrk.readStringValue("STARTUP-PATH");
				} finally {
					wrk.close();
				}
			} catch(e) {}
			
			// create nsIFile from path in registry
			if(path) {
				try {
					var startupFolder = Components.classes["@mozilla.org/file/local;1"].
						createInstance(Components.interfaces.nsILocalFile);
					startupFolder.initWithPath(path);
					startupFolders.push(startupFolder);
				} catch(e) {
					addDefaultStartupFolder = true;
				}
			} else {
				try {
					wrk.open(Components.interfaces.nsIWindowsRegKey.ROOT_KEY_CURRENT_USER,
						"Software\\Microsoft\\Office\\"+i+".0\\Common\\General",
						Components.interfaces.nsIWindowsRegKey.ACCESS_READ);
					try {
						var startup = wrk.readStringValue("Startup");
						var startupFolder = appData.clone().QueryInterface(Components.interfaces.nsILocalFile);
						startupFolder.appendRelativePath("Microsoft\\Word\\"+startup);
						startupFolders.push(startupFolder);
					} finally {
						wrk.close();
					}
				} catch(e) {
					addDefaultStartupFolder = true;
				}
			}
		}
		
		if(startupFolders.length == 0 || addDefaultStartupFolder) {
			// if not in the registry, append Microsoft/Word/Startup to %AppData% (default location)
			var startupFolder = appData.clone().QueryInterface(Components.interfaces.nsILocalFile);
			startupFolder.appendRelativePath("Microsoft\\Word\\Startup");
			startupFolders.push(startupFolder);
		}
		
		for each(var startupFolder in startupFolders) {
			var oldDot = startupFolder.clone().QueryInterface(Components.interfaces.nsILocalFile);
			oldDot.append("Zotero.dot");
			if(oldDot.exists()) oldDot.remove(false);
			
			// copy Zotero.dot file to Word Startup folder
			dot.copyTo(startupFolder, "Zotero.dot");
		}
	} catch(e) {
		var prompts = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
			.getService(Components.interfaces.nsIPromptService)
			.alert(null, 'Zotero Word Integration Error',
			'Zotero Word Integration could not complete installation because an error occurred. Please ensure that Word is closed, and then restart Firefox.');
		throw e;
	}
}

var ext = Components.classes['@mozilla.org/extensions/manager;1']
   .getService(Components.interfaces.nsIExtensionManager).getItemForID(ZOTEROWINWORDINTEGRATION_ID);
var ZoteroWinWordIntegration_prefService = Components.classes["@mozilla.org/preferences-service;1"].
	getService(Components.interfaces.nsIPrefBranch);
if(ZoteroWinWordIntegration_prefService.getCharPref(ZOTEROWINWORDINTEGRATION_PREF) != ext.version) {
	ZoteroWinWordIntegration_firstRun();
	ZoteroWinWordIntegration_prefService.setCharPref(ZOTEROWINWORDINTEGRATION_PREF, ext.version);
}