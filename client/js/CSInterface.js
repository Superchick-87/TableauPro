/**************************************************************************************************
 *
 * ADOBE SYSTEMS INCORPORATED
 * Copyright 2013 Adobe Systems Incorporated
 * All Rights Reserved.
 *
 * NOTICE:  Adobe permits you to use, modify, and distribute this file in accordance with the
 * terms of the Adobe license agreement accompanying it.  If you have received this file from a
 * source other than Adobe, then your use, modification, or distribution of it requires the prior
 * written permission of Adobe.
 *
 **************************************************************************************************/

/**
 * CSInterface - v9.0.0
 */

/**
 * @class CSInterface
 * This is the entry point to the CEP Extensibility infrastructure.
 * Instantiate this class and use the methods to interact with the host application.
 */
var CSInterface = function() {
};

/**
 * User can provide their own implementation of window.JSON
 */
if (typeof JSON == "undefined") {
    if (window.JSON) {
        JSON = window.JSON;
    } else if (window.cep && window.cep.util) {
        JSON = window.cep.util;
    }
}

/**
 * @return the host environment data object.
 */
CSInterface.prototype.getHostEnvironment = function() {
    var hostEnv = window.cep.fs.readFile(decodeURI(window.cep.extension.getHostEnvironmentFile()));
    var hostEnvObj = JSON.parse(hostEnv);
    
    // Add appId to the host environment object to support older extensions
    if (!hostEnvObj.appId) {
        hostEnvObj.appId = hostEnvObj.appId;
    }
    
    return hostEnvObj;
};

/**
 * Closes this extension.
 */
CSInterface.prototype.closeExtension = function() {
    window.cep.extension.closeExtension();
};

/**
 * @return the system path of the extension.
 */
CSInterface.prototype.getSystemPath = function(pathType) {
    var path = decodeURI(window.cep.fs.readFile(window.cep.extension.getSystemPath(pathType)));
    var originPath = path;
    if(navigator.platform == "Win32") {
        path = path.substring(8);
    } else if(navigator.platform == "MacIntel") {
        path = path.substring(7);
    }
    return path;
};

/**
 * Evaluates a JavaScript script, which can use the JavaScript DOM
 * of the host application.
 *
 * @param script    The JavaScript script.
 * @param callback  Optional. A callback function that receives the result of execution.
 * If execution fails, the callback function receives the error message
 * EvalScript_ErrMessage.
 */
CSInterface.prototype.evalScript = function(script, callback) {
    if(callback === null || callback === undefined) {
        callback = function(result){};
    }
    window.cep.process.evalScript(script, callback);
};

/**
 * @return the App skin information.
 */
CSInterface.prototype.getHostCapabilities = function() {
    var hostCapabilities = JSON.parse(window.cep.fs.readFile(decodeURI(window.cep.extension.getHostCapabilitiesFile())));
    return hostCapabilities;
};

/**
 * Opens a page in the default system browser.
 *
 * @param url  The URL of the page/file to open, or the email address.
 * Must use http://, https://, file:// or mailto: schema.
 */
CSInterface.prototype.openURLInDefaultBrowser = function(url) {
    return cep.util.openURLInDefaultBrowser(url);
};

/**
 * @return the extension ID.
 */
CSInterface.prototype.getExtensionID = function() {
    return window.cep.extension.getExtensionId();
};

/**
 * @return the scale factor of screen.
 */
CSInterface.prototype.getScaleFactor = function() {
    return window.cep.util.getScaleFactor();
};

/**
 * @return the scale factor of screen.
 */
CSInterface.prototype.setScaleFactor = function() {
    return window.cep.util.setScaleFactor();
};

/**
 * @return current OS information.
 */
CSInterface.prototype.getOSInformation = function() {
    var userAgent = navigator.userAgent;

    if ((navigator.platform == "Win32") || (navigator.platform == "Windows")) {
        var winData = "Windows";
        var winVer = /Windows NT([^;]+)/.exec(userAgent);
        if (winVer && winVer.length > 1) {
            winData += " " + winVer[1];
        }
        return winData;
    } else if (navigator.platform == "MacIntel") {
        var macData = "Mac OS X";
        var macVer = /Mac OS X([^)]+)/.exec(userAgent);
        if (macVer && macVer.length > 1) {
            macData += " " + macVer[1];
        }
        return macData;
    }

    return "Unknown";
};

/**
 * Register an interest in a specific event, and request that the
 * callback function be executed when the event occurs.
 *
 * @param type      The name of the event.
 * @param listener  The function to be executed when the event occurs.
 * @param obj       Optional. An object containing the scope for the listener.
 */
CSInterface.prototype.addEventListener = function(type, listener, obj) {
    window.cep.extension.addEventListener(type, listener, obj);
};

/**
 * Removes the interest in a specific event.
 *
 * @param type      The name of the event.
 * @param listener  The function that was previously registered.
 * @param obj       Optional. An object containing the scope for the listener.
 */
CSInterface.prototype.removeEventListener = function(type, listener, obj) {
    window.cep.extension.removeEventListener(type, listener, obj);
};

/**
 * Trigget a CEP event.
 *
 * @param event A CSEvent object.
 */
CSInterface.prototype.dispatchEvent = function(event) {
    if (typeof event.data == "object") {
        event.data = JSON.stringify(event.data);
    }

    window.cep.extension.dispatchEvent(event);
};

/**
 * CSEvent - v9.0.0
 *
 * @class CSEvent
 * @param type          The name of the event.
 * @param scope         The scope of event, can be "GLOBAL" or "APPLICATION".
 * @param appId         The unique identifier of the application that generated the event.
 * @param extensionId   The unique identifier of the extension that generated the event.
 */
var CSEvent = function(type, scope, appId, extensionId) {
    this.type = type;
    this.scope = scope;
    this.appId = appId;
    this.extensionId = extensionId;
};

/**
 * SystemPath - v9.0.0
 */
var SystemPath = {
    USER_DATA: "userData",
    COMMON_FILES: "commonFiles",
    MY_DOCUMENTS: "myDocuments",
    APPLICATION: "application",
    EXTENSION: "extension",
    HOST_APPLICATION: "hostApplication"
};


/**
 * ColorType - v9.0.0
 */
var ColorType = {
    RGB: "rgb",
    GRADIENT: "gradient",
    NONE: "none"
};

/**
 * RGBColor - v9.0.0
 */
var RGBColor = function(red, green, blue, alpha) {
    this.type = ColorType.RGB;
    this.red = red;
    this.green = green;
    this.blue = blue;
    this.alpha = alpha;
};

/**
 * Direction - v9.0.0
 */
var Direction = {
    x: "x",
    y: "y"
};

/**
 * GradientStop - v9.0.0
 */
var GradientStop = function(offset, rgbColor) {
    this.offset = offset;
    this.rgbColor = rgbColor;
};

/**
 * GradientColor - v9.0.0
 */
var GradientColor = function(type, direction, arrGradientStop) {
    this.type = ColorType.GRADIENT;
    this.gradientType = type;
    this.direction = direction;
    this.arrGradientStop = arrGradientStop;
};

/**
 * UIColor - v9.0.0
 */
var UIColor = function(type, antialiasLevel, color) {
    this.type = type;
    this.antialiasLevel = antialiasLevel;
    this.color = color;
};

/**
 * AppSkinInfo - v9.0.0
 */
var AppSkinInfo = function(baseFontFamily, baseFontSize, appBarBackgroundColor, panelBackgroundColor, appBarBackgroundColorSR, panelBackgroundColorSR) {
    this.baseFontFamily = baseFontFamily;
    this.baseFontSize = baseFontSize;
    this.appBarBackgroundColor = appBarBackgroundColor;
    this.panelBackgroundColor = panelBackgroundColor;
    this.appBarBackgroundColorSR = appBarBackgroundColorSR;
    this.panelBackgroundColorSR = panelBackgroundColorSR;
};

/**
 * HostEnvironment - v9.0.0
 */
var HostEnvironment = function(appName, appVersion, appLocale, appUILocale, appId, isAppOnline, appSkinInfo) {
    this.appName = appName;
    this.appVersion = appVersion;
    this.appLocale = appLocale;
    this.appUILocale = appUILocale;
    this.appId = appId;
    this.isAppOnline = isAppOnline;
    this.appSkinInfo = appSkinInfo;
};

/**
 * HostCapabilities - v9.0.0
 */
var HostCapabilities = function(EXTENDED_PANEL_MENU, EXTENDED_PANEL_ICONS, DELEGATE_APE_ENGINE, SUPPORT_HTML_EXTENSIONS, DISABLE_FLASH_EXTENSIONS) {
    this.EXTENDED_PANEL_MENU = EXTENDED_PANEL_MENU;
    this.EXTENDED_PANEL_ICONS = EXTENDED_PANEL_ICONS;
    this.DELEGATE_APE_ENGINE = DELEGATE_APE_ENGINE;
    this.SUPPORT_HTML_EXTENSIONS = SUPPORT_HTML_EXTENSIONS;
    this.DISABLE_FLASH_EXTENSIONS = DISABLE_FLASH_EXTENSIONS;
};

/**
 * ApiVersion - v9.0.0
 */
var ApiVersion = function(major, minor, micro) {
    this.major = major;
    this.minor = minor;
    this.micro = micro;
};

/**
 * MenuItemStatus - v9.0.0
 */
var MenuItemStatus = function(menuItemLabel, enabled, checked) {
    this.menuItemLabel = menuItemLabel;
    this.enabled = enabled;
    this.checked = checked;
};

/**
 * ContextMenuItemStatus - v9.0.0
 */
var ContextMenuItemStatus = function(menuItemLabel, enabled, checked) {
    this.menuItemLabel = menuItemLabel;
    this.enabled = enabled;
    this.checked = checked;
};