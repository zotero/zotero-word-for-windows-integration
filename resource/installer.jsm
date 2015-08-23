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

var EXPORTED_SYMBOLS = ["Installer"];
var Zotero = Components.classes["@zotero.org/Zotero;1"].getService(Components.interfaces.nsISupports).wrappedJSObject;
var ZoteroPluginInstaller = Components.utils.import("resource://zotero-winword-integration/installer_common.jsm").ZoteroPluginInstaller;
var Installer = function(failSilently, force) {
	return new ZoteroPluginInstaller(Plugin,
		failSilently !== undefined ? failSilently : Zotero.isStandalone,
		force);
}

var Plugin = new function() {
	this.EXTENSION_STRING = "Zotero Word for Windows Integration";
	this.EXTENSION_ID = "zoteroWinWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroWinWordIntegration.";
	this.EXTENSION_DIR = "zotero-winword-integration";
	this.APP = 'Microsoft Word';
	
	this.REQUIRED_ADDONS = [{
		name: "Zotero",
		url: "zotero.org",
		id: "zotero@chnm.gmu.edu",
		minVersion: "4.0.27.SOURCE",
		required: true
	}, {
		name: "Zotero LibreOffice Integration",
		url: "zotero.org/support/word_processor_plugin_installation",
		id: "zoteroOpenOfficeIntegration@zotero.org",
		minVersion: "3.5b2.SVN",
		required: false
	}];
	this.LAST_INSTALLED_FILE_UPDATE = "3.2.0";

	var zoteroPluginInstaller;
	
	this.verifyNotCorrupt = function(zpi) {}

	this.install = function(zpi) {
		// find Word Startup folders (see http://support.microsoft.com/kb/210860)
		var appData = Components.classes["@mozilla.org/file/directory_service;1"]
				.getService(Components.interfaces.nsIProperties)
				.get("AppData", Components.interfaces.nsILocalFile);

		// first check the registry for a custom startup folder
		var startupFolders = [];
		var addDefaultStartupFolder = false;
		var wrk = Components.classes["@mozilla.org/windows-registry-key;1"]
			.createInstance(Components.interfaces.nsIWindowsRegKey);
		var ZoteroDot = "Zotero.dot";
		for(var i=9; i<=14; i++) {
			var path = null;
			try {
				wrk.open(Components.interfaces.nsIWindowsRegKey.ROOT_KEY_CURRENT_USER,
					"Software\\Microsoft\\Office\\"+i+".0\\Word\\Options",
					Components.interfaces.nsIWindowsRegKey.ACCESS_READ);
				// If registry exists for Word 2007 or above, use Zotero.dotm
				if (i >= 12) {
					ZoteroDot = "Zotero.dotm";
				}
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

		// get Zotero.dotm file
		var dot = zpi.getAddonPath(this.EXTENSION_ID);
		dot.append("install");
		dot.append(ZoteroDot);

		if(startupFolders.length == 0 || addDefaultStartupFolder) {
			// if not in the registry, append Microsoft/Word/Startup to %AppData% (default location)
			var startupFolder = appData.clone().QueryInterface(Components.interfaces.nsILocalFile);
			startupFolder.appendRelativePath("Microsoft\\Word\\Startup");
			startupFolders.push(startupFolder);
		}

		for each(var startupFolder in startupFolders) {
			var oldDot = startupFolder.clone().QueryInterface(Components.interfaces.nsILocalFile);
			if(oldDot.equals(dot)) continue;
			oldDot.append(ZoteroDot);
			if(oldDot.exists()) {
				try {
					oldDot.remove(false);
				} catch(e) {
					throw "Could not remove "+oldDot.path;
				}
			}

			// copy Zotero.dotm file to Word Startup folder
			dot.copyTo(startupFolder, ZoteroDot);
		}

		zpi.success();
	}
}
