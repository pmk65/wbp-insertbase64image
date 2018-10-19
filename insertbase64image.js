/**
 * Insert Base64 Image Plugin
 *
 * Insert Base64 encoded image data in editor
 * Supported image filetypes: .gif, .png, .jpg, .svg and .ico
 *
 * @category  WeBuilder Plugin
 * @package   Inset Base64 Image
 * @author    Peter Klein <pmk@io.dk>
 * @copyright 2017
 * @license   http://www.freebsd.org/copyright/license.html  BSD License
 * @version   2.0
 */

/**
 * [CLASS/FUNCTION INDEX of SCRIPT]
 *
 *     35   function InsertBase64Image()
 *    113   function Base64EncodeImage(img)
 *    145   function SelectImageFile()
 *    165   function ConvertBytes(bytes)
 *    180   function Power(num, count)
 *    192   function OnInstalled()
 *
 * TOTAL FUNCTIONS: 6
 * (This index is automatically created/updated by the WeBuilder plugin "DocBlock Comments")
 *
 */

/**
 * Insert Base64 encoded image in editor
 *
 * @return void
 */
function InsertBase64Image() {

	var imageFile = SelectImageFile();
	if (imageFile == "") return;

	var path = ExtractFilePath(imageFile);
	var name = ExtractFileName(imageFile);
	var ext = Lowercase(ExtractFileExt(imageFile));
	var width = "", height = "", cr = chr(13);

	var oShell = CreateOleObject("Shell.Application");
	var oFile = oShell.NameSpace(path).ParseName(name);

	var size = StrToInt(VarToStr(oFile.ExtendedProperty("Size")));

	// Show confirm dialog if image file is larger than 10KB
	if ((size > 10240)) {
		if (Confirm("The selected image is quite large (" + ConvertBytes(size) + ")" + cr + "Are you really sure you want to insert it?") == false) return;
	}

	if (ext != ".svg") {

		// Get image dimensions and remove unwanted chars from result
		var dim = RegexReplace(VarToStr(oFile.ExtendedProperty("Dimensions")), "[^\\dx ]", "", true);

		// Not an image
		if (dim == "") return;

		width = RegexMatch(dim, "^\\d+", true);
		height = RegexMatch(dim, "\\d+$", true);

	}

	var base64 = Base64EncodeImage(imageFile);
	if (base64 == "") return;

	if (Script.ReadSetting("Remove linefeeds from Base64 data", "1") == "1") {
		base64 = RegexReplace(base64, "\\n", "", true);
	}

	var res = "\"" + base64 + "\"" + cr;

	// Wrap base64 output based on editor codetype
	if (Script.ReadSetting("Wrap result based on codetype", "1") == "1") {
		switch (Document.CurrentCodeType) {
			case ltCSS: res = "background-image: url('" + base64 + "'); /* " + name + " */" + cr +
								"/*" + cr +
								"width: " + width + "px;" + cr +
								"height: " + height + "px;" + cr +
								"*/" + cr;

			case ltHTML: res = "<img src=\"" + base64 + "\" width=\"" + width + "\" height=\"" + height + "\" alt=\"" + name + "\" />" + cr;

			case ltJScript: res = "// " + name +", width: " + width + "px, height: " + height + "px" + cr +
								  "var imageData = '" + RegexReplace(base64, "\\n", "", true) + "';" + cr;

			case ltPHP: res = "// " + name +", width: " + width + "px, height: " + height + "px" + cr +
								"$imageData = '" + base64 + "';" + cr;

			case ltASP: res = "/* " + name +", width: " + width + "px, height: " + height + "px" + " */" + cr +
								"imageData = \"" + base64 + "\""  + cr;

			case ltXML, ltWML: res = "<image encoding=\"base64\">" + RegexReplace(base64, "^data:image\\/.*;base64,", "", true) + "</image>" + cr;
		}
	}

	// Insert result at cursor position
	Editor.SelText = res;
	Editor.Focus;
}

/**
 * Encode imagedata into base64 format
 *
 * @param  string   img  full path to imagefile
 *
 * @return string   base64 encoded data
 */
function Base64EncodeImage(img) {

	var oXML = CreateOleObject("MSXML2.DOMDocument.6.0");
	var oNode = oXML.CreateElement("base64");
	oNode.dataType = "bin.base64";

	var oStream = CreateOleObject("ADODB.Stream");
	oStream.Type = 1;
	oStream.Open;
	oStream.LoadFromFile(img);
	oNode.nodeTypedValue = oStream.Read();
	var base64 = oNode.Text;
	oStream.Close;

	// Mime type detection based on first 4 bytes of base64 encoded data
	var mime = "", seq = Copy(base64, 0, 4);
	if (seq == "iVBO")      mime = "png";
	else if (seq == "R0lG") mime = "gif";
	else if (seq == "/9j/") mime = "jpeg";
	else if (seq == "PD94") mime = "svg+xml";
	else if (seq == "AAAB") mime = "x-icon";

	if (mime == "") return "";

	return "data:image/" + mime + ";base64," + base64;
}

/**
 * Bring up dialog to select image file.
 *
 * @return string   path/filename name of image file selected
 */
function SelectImageFile() {

	var imageFile = "", oDialog = new TOpenDialog(WeBuilder);
	oDialog.InitialDir = ExtractFilePath(Document.Filename);
	oDialog.Title = "Select image file";
	oDialog.Filter = "Web image files only|*.gif;*.png;*.jpg;*.svg;*.ico";
	oDialog.Options = ofFileMustExist + ofNoDereferenceLinks;
	if (oDialog.Execute) imageFile = oDialog.FileName;
	delete oDialog;

	return imageFile;
}

/**
 * Convert bytes value into short notation
 *
 * @param  integer   bytes
 *
 * @return string
 */
function ConvertBytes(bytes) {
	var i = 0, units = ["Bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"];
	while (bytes > Power(1024, i + 1)) i++;
	return FormatFloat("###0.##", Round(bytes / Power(1024, i))) + " " + units[i];
}

/**
 * Math Power function
 * (Positive integer values only)
 *
 * @param  integer   num
 * @param  integer   count
 *
 * @return integer
 */
function Power(num, count) {
	if (count == 0) return 1;
	var t = Power(num, count/2);
	if (count%2 == 0) return t*t;
	else return num*t*t;
}

/**
 * Show info when plugin is installed
 *
 * @return void
 */
function OnInstalled() {
	alert("Insert Base64 Image 2.0 by Peter Klein installed sucessfully!");
}

Script.ConnectSignal("installed", "OnInstalled");
var bmp = new TBitmap, act;
LoadFileToBitmap(Script.Path + "base64_icon.png", bmp);
act = Script.RegisterDocumentAction("", "Insert Base64 Image", "", "InsertBase64Image");
Actions.SetIcon(act, bmp);
delete bmp;