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

var ZoteroWinWordIntegration = new function() {
	this.EXTENSION_STRING = "Zotero WinWord Integration";
	this.EXTENSION_ID = "zoteroWinWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroWinWordIntegration.";
	this.EXTENSION_DIR = "zotero-winword-integration";
	this.APP = 'Microsoft Word';
	
	this.REQUIRED_ADDONS = [{
		name: "Zotero",
		url: "zotero.org",
		id: "zotero@chnm.gmu.edu",
		minVersion: "2.1a1.SVN"
	}];
	
	var zoteroPluginInstaller;
	
	this.verifyNotCorrupt = function(zpi) {}
	
	this.install = function(zpi) {
		// get Zotero.dot file
		var dot = zpi.getAddonPath(this.EXTENSION_ID);
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
		for(var i=9; i<=14; i++) {
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
		
		zpi.success();
	}
}