/**
 * Script for David
 *
 */

#include "xlsx.extendscript.js";

var script = function() {
    var total=0;
    var complete=0;
    var percentage=0;
    var mypercentage;
    var wantstop=false;
    function cTID(s) {
        return app.charIDToTypeID(s);
    };

    String.prototype.trim2 = function() {
        return this.replace(/^\s+|\s+$/g, '');
    };

    function indexOf(arr, obj) {
        var i, j;
        for (i = -1, j = arr.length; ++i < j;) {
            if (arr[i] == obj) {
                return i;
            }
        }
        return -1;
    }


    var keyword_image = 1;
    var keyword_text = 2;


   var originalRulerUnits = preferences.rulerUnits;
    preferences.rulerUnits = Units.PIXELS;
    var startDisplayDialogs = app.displayDialogs;
    app.displayDialogs = DialogModes.NO;

    var get_coordinate = '';
    var get_device_model = '';
   
    var dquality = 12; // compress quality from 1 to 12

    var psource = File($.fileName).path;
    var jpegSaveOptions = new JPEGSaveOptions();
    jpegSaveOptions.quality = dquality;
    jpegSaveOptions.embedICCProfile = false;
    jpegSaveOptions.embedColorProfile = false;     

    if (ExternalObject.AdobeXMPScript != undefined) ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");
    var csvR;
    var targFR;
    var targFE;
    var targSR;
    var targSE;
    var randomchk = true;
    

    try {
        // dialog
        var progressWindow = createProgressWindow("Photoshop BULK Content Creator", false, true);
        progressWindow.isDone = false;
        progressWindow.onCancel = function() {
            this.isDone = true;
            return true; 
        }
        progressWindow.updateProgress(0);
        while (progressWindow.wait && !progressWindow.isDone) {
            progressWindow.update();
            if (!progressWindow.visible) return;
            var myKeyState = ScriptUI.environment.keyboardState;
            if (myKeyState.keyName !== undefined) {
                if (myKeyState.keyName.toLowerCase() == "escape") {
                    return;
                }
            }
        }
        if (progressWindow.isDone) return;

            try {
                csvR = progressWindow.csvR.text !== "" ? new File(progressWindow.csvR.text) : null;
                var dd1 = null;
                var dd2 = null;

                if (csvR !== null && csvR.exists) {
                    dd1 = readdat(csvR);
                }
            } catch (e) {
                alert("#5 " + e + e.line);
            } finally {
                try {
                    docRef1.close(SaveOptions.DONOTSAVECHANGES);
                } catch (_) {}
                try {
                    docRef2.close(SaveOptions.DONOTSAVECHANGES);
                } catch (_) {}
            }
    } catch (e) {
        alert("#4 " + e + e.line);
    } finally {
        progressWindow.close();
    }

    function addMetaData(file,lat,lng){    
        // takes a file path, lat value, and long value  
        // adds to the metadata  
      
        if ( !ExternalObject.AdobeXMPScript ) ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');    
            var xmpf = new XMPFile( File(file).fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_UPDATE );    
            var xmp = xmpf.getXMP();    
            xmp.deleteProperty(XMPConst.NS_EXIF, "GPSLatitude");  
            xmp.setProperty(XMPConst.NS_EXIF, "GPSLatitude", lat);   
            xmp.deleteProperty(XMPConst.NS_EXIF, "GPSLongitude");  
            xmp.setProperty(XMPConst.NS_EXIF, "GPSLongitude", lng);   
            if (xmpf.canPutXMP( xmp )) {    
            xmpf.putXMP( xmp );    
            }    
            xmpf.closeFile( XMPConst.CLOSE_UPDATE_SAFELY );    
    }
    

    /**
     * read dat file
     * @param   {File} fcsv
     * @returns {Object|Boolean}
     */
    function readdat(fcsv) {
        var res = {};
        var fop;
        try {
            //fcsv.encoding = 'UTF-8';
            fop = fcsv.open("r:");
            var sss;
            res.dat = [];
            var line = 0;
            for (; !fcsv.eof;) {
                sss = fcsv.readln();
                res.dat[line] = sss.trim2();
                line++;
            }
        } catch (ex) {
            alert("#2 " + ex + ex.line);
        } finally {
            if (fop) fcsv.close();
        }
        return res;
    }

    function csvToArray(text) {
        var ret = [''],
            i = 0,
            p = '',
            s = true;
        for (var n = 0, l; n < text.length; n++) {
            l = text.charAt(n);
            if ('"' === l) {
                s = !s;
                if ('"' === p) {
                    ret[i] += '"';
                    l = '-';
                } else if ('' === p)
                    l = '-';
            } else if (s && ';' === l)
                l = ret[++i] = '';
            else
                ret[i] += l;
            p = l;
        }
        return ret;
    }

    /* createProgressWindow
       title     the window title
       parent    the parent ScriptUI window (opt)
       useCancel flag for having a Cancel button (opt)
    
       onCancel  This method will be called when the Cancel button is pressed.
                 This method should return 'true' to close the progress window
    */
    function createProgressWindow(title, parent, useCancel) {
        var win = new Window('palette', title);
        var grpm = win.add("group");
        grpm.orientation = "column";
        grpm.alignment = ["fill", "top"];
        g1 = win.add("group {orientation: 'row', alignChildren: ['fill', 'fill'], margins: 0}");

        g1.add("statictext", [0, 0, 100, 18], "percentage", "alignment:center");
        mypercentage = g1.add("statictext", [0, 0, 100, 18], "0 %", "alignment:center");

        g1 = win.add("group {orientation: 'row', alignChildren: ['fill', 'fill'], margins: 0}");
        g1.add("statictext", [100, 50, 220, 18], "File Path:", "alignment:right");
        win.csvR = g1.add("edittext", undefined, "");
        win.csvR.minimumSize.width = 450;
        var btn1 = g1.add("button", undefined, "BROWSE");
        btn1.onClick = function() {
            var fcsv = File.openDialog("Select file", "CSV:*.csv,All files:*.*", psource);
            if (fcsv === null) {
                return;
            } else {
                win.csvR.text = File.decode(fcsv);
            }
        };
        btn1.maximumSize.width = 60;
        btn1.alignment = ["right", "fill"];
        g1 = win.add("group {orientation: 'row', alignChildren: ['fill', 'fill'], margins: 0}");
        g1.alignment = ["fill", "fill"];
        win.wait = true;
        win.start = g1.add('button', [0, 0, 220, 20], 'START');
        win.start.alignment = ["left", "fill"];
        win.start.onClick = function() {
            win.wait = false;
            callme();
        };

        //  win.parent = undefined;

        if (parent) {
            if (parent instanceof Window) {
                win.parent = parent;
            } else if (useCancel == undefined) {
                useCancel = parent;
            }
        }

        if (useCancel) {
            win.cancel = g1.add('button', [0, 0, 220, 20], 'Cancel');
            win.cancel.alignment = ["left", "fill"];
            win.cancel.onClick = function() {
                try {
                    if (win.onCancel) {
                        var rc = win.onCancel();
                        if (rc || rc == undefined) {
                            wantstop=true;
                            win.close();
                        }
                    } else {
                        win.close();
                    }
                } catch (e) {
                    alert("#3 " + e + e.line);
                }
            }
        }

        win.progresstext = g1.add("statictext", undefined, "", "alignment:left");
        win.progresstext.alignment = ["fill", "middle"];


        win.updateProgress = function(val) {
            var win = this;
            win.show();
            win.hide();
            win.show();
            win.update();
            app.refresh();
        }

        win.center(win.parent);

        return win;
    }


	function addLatLong(str){      
		if (str) {      
			var latLong = str.split(",");      
			if (latLong[0] && latLong[1]){      
		  
				/// LATITUDE ///////      
				var lat = latLong[0];      
				var latDir = ''; //N or S      
				if (lat>0){latDir = 'N';} else {latDir = 'S';} // get North or South      
		  
				lat = convertToDM(lat) + latDir; // final lat value      
													   
				/// LONGITUDE //////      
				var lng = latLong[1];      
				var lngDir = ''; // E or W      
				if (lng>0){lngDir = 'E';} else {lngDir = 'W';} // get East or West      
		  
				lng = convertToDM(lng) + lngDir; // final lng value      
					  
                  return lat + '|' + lng;
					  
			} else {      
				return '';      
			}      
		}      
	}
      
	function convertToDM(str){      
		// takes DD.dddddd and returns DD MM.mmmm      
			  
		var split = str.split(".");      
		var D = split[0]; // get degrees      
		D = Math.abs(D); // absolute value to remove negative sign      
		var M = '.' + split[1]; // get minutes      
		M = M*60; // convert      
		M = M.toFixed(4); // round to Adobe requirements      
		  
		str = D + "," + M; // final value          
		return str;      
	}  
    
    function removeMetadata(file_path){  
        var file = File(file_path);  
        if ( !ExternalObject.AdobeXMPScript ) ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');  
            
            var xmpf = new XMPFile( File(file).fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_UPDATE );  
            var xmp = xmpf.getXMP();  
            
            XMPUtils.removeProperties(xmp, "", "", XMPConst.REMOVE_ALL_PROPERTIES);  
        
            if (get_coordinate != '')
            {
                var refine_coordinate = get_coordinate.split(','); 
            
                var refine_lat = refine_coordinate[0];
                refine_lat = refine_lat.substring(0, refine_lat.length-3) + (Math.random().toString().replace('0.', '').substring(0, 3)).toString();
            
                var refine_lng = refine_coordinate[1];
                refine_lng = refine_lng.substring(0, refine_lng.length-3) + (Math.random().toString().replace('0.', '').substring(0, 3)).toString();
            
                var get_final_coordinate = addLatLong(refine_lat + ', ' + refine_lng);
            
                if (get_final_coordinate != '')
                {
                    var final_coordinate = get_final_coordinate.split('|');
                    xmp.setProperty(XMPConst.NS_EXIF, "GPSLatitude", final_coordinate[0]);  
                    xmp.setProperty(XMPConst.NS_EXIF, "GPSLongitude", final_coordinate[1]);  
                }
            }
            
             if (get_device_model != '')
             {
                 
                 xmp.setProperty(XMPConst.NS_CAMERA_RAW, "Model", get_device_model);  
             }
        
        if (xmpf.canPutXMP( xmp )) {  
            xmpf.putXMP( xmp );  
        }  
        xmpf.closeFile( XMPConst.CLOSE_UPDATE_SAFELY );  
    };


function callme(){
   var dd1=null;
   var data = getExcelData(progressWindow.csvR.text);
   var rule = data[0];//this value will show the head so this will describe what should do ex: templete. image. text. save....
   var index = 0;
   var word ;
   var temp ;
   var sclreduct = 1;//if inupt 
   var targetFolder = null;
   var importFolder = null;
   
   var total_info = 0;
   var isfirst = true;
   var temppath = data[1][0];

   /**
    * sclreduct = number(reduct)/100.0
    * 
    */
    for(var x = 1; x<data.length;x++){
        if(data[x][0]!=undefined)
        total++;
    }
  



   try{
   for(; index<data.length-1;index=index+1){
    //    alert(index);
      word = data[index+1];
    //   for(var subindex=0; subindex < word.length;subindex++){
        //  if(word[index] == undefined)continue;
         
        //  getresult(word, index, rule, index);
        
        for(var subindex=0; subindex < word.length;subindex++){
            if(temppath!=word[0]) {isfirst = true;
                }
            if(rule[subindex].indexOf("Template") !=-1){
                var docRef1lays = [];
                temppath = word[subindex];
                temp = app.open(new File(word[subindex]));
                temp.resizeImage(temp.width * sclreduct, temp.height * sclreduct);
                if(temp !== null){
                    app.activeDocument = temp;
                    setBottomTextFields(temp, rule, docRef1lays);
                    if(isfirst){
                    createNamedSnapshot("stp0");
                    isfirst=false;
                    }
                   
        
                }
            }


            if(((rule[subindex].indexOf("Save")) !=-1)&&!wantstop){
                
                targetFolder =word[subindex];
                make1(temp,word, index, docRef1lays, targetFolder, index,rule);
                
            }

        }
        if(wantstop){
            alert("Image creation canceled");
            break;
        }
        
        
        // if((rule[subindex].indexOf("Image")) !=-1){
        //     importFolder = new Folder(word[subindex]);
            
        // }
        // if(targetFolder !== null && !targetFolder.exist) targetFolder.create();
        
        
    //   }
      
      
   }
}
catch(E){
    alert("some error :"+e+"From"+e.line+"please ask to Rolland");
}
   
}


function replaceContents (newFile) {  
   // =======================================================  
   var idplacedLayerReplaceContents = stringIDToTypeID( "placedLayerReplaceContents" );  
       var desc3 = new ActionDescriptor();  
       var idnull = charIDToTypeID( "null" );  
       desc3.putPath( idnull, new File( newFile ) );  
       var idPgNm = charIDToTypeID( "PgNm" );  
       desc3.putInteger( idPgNm, 1 );  
   executeAction( idplacedLayerReplaceContents, desc3, DialogModes.NO );  
   return app.activeDocument.activeLayer  
};


function getExcelData ( path ){

   try{
   /* Read file from disk */
   var workbook = XLSX.readFile(path , {cellDates:true});
   /* Display first worksheet */
   var first_sheet_name = workbook.SheetNames[0], first_worksheet = workbook.Sheets[first_sheet_name];
   var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});
   return data;
   }
   catch(E){
       alert("this error is"+ E+"in"+E.line);
   }
   
}

function meta_comments(fld, file, comments) {
    var paramsfile = new File(fld + "/generate.params");
    paramsfile.open("w:");
    paramsfile.write(File.decode(file.fsName) + '\t' + comments);
    paramsfile.close();

    var exfile;
    if ($.os.match(/windows/i)) {
        exfile = new File(fld + "/meta_comments.exe");
        if (exfile.exists) {
            exfile.execute();
        }
    } else {
        exfile = new File(fld + "/meta_comments.app");
        if (exfile.exists) {
            exfile.execute();
        }
    }
}

function make1(docRef, dd, row, docReflays, targ, foldernum,rule) {
    // alert("hey :"+docRef+"dd :"+dd+"row :"+row+"doc :"+docReflays+"fodernum :"+foldernum);
    
    try {
        // app.activeDocument = docRef;
        //docRef.activeHistoryState = savedState;
        revertNamedSnapshot("stp0");

        // var word = '' + dd.dat[row];
        var layernum = 0;
        var words = dd
        // alert(docReflays);
        for (var ilay = 0; ilay < docReflays.length; ilay++) {
            var layerRef = docReflays[ilay].lay;
            docRef.activeLayer = layerRef;
            //if (docReflays[ilay].type == "txt") {
            if (docReflays[ilay].lay.kind == LayerKind.TEXT) {
                //var layerRef = docRef.artLayers.getByName( 'Text' );
                for(var i = 0;i<rule.length;i++){
                    if(rule[i].substring(rule[i].indexOf("#")+1,rule[i].length)==docReflays[ilay].lay.name)
                    var contents = words[i];
                }
                
                // var contents = words[docReflays[ilay].col].trim2();
                if (contents !== '') {
                    //setbound(undefined, docRef.height);
                    var currentFontSize = docReflays[ilay].size * 0.1;
                    fontSize(currentFontSize, docReflays[ilay].lea);
                    layerRef.textItem.contents = contents; //words[1];
                    //var textHeight = layerRef.bounds[3] - layerRef.bounds[1];
                    //var contents_ws = contents.replace(/\s/g, '');
                    //var boundHeight = layerRef.textItem.height;
                    //fitText(docReflays[ilay].gr300, currentFontSize, layerRef, contents_ws.length);
                    fitText(docReflays[ilay].gr300, currentFontSize, docReflays[ilay].lea, layerRef);

                    //var text_shift = (docReflays[ilay].wwheight - (layerRef.bounds[3] - layerRef.bounds[1])) * 0.5;
                    //setbound(-text_shift);
                    //setbound(-100);
                }
            } else if (docReflays[ilay].lay.kind == LayerKind.SMARTOBJECT) {
                //var theLayer = app.activeDocument.artLayers.getByName( 'Background' );

                    for(var i = 0;i<rule.length;i++){
                        if(rule[i].substring(rule[i].indexOf("#")+1,rule[i].length)==docReflays[ilay].lay.name)
                        var nameinfile = words[i];
                    }
                // var nameinfile = words[docReflays[ilay].col].trim2();
                if (nameinfile !== '') {
                    var infile = new File(nameinfile); //words[0]);
                    if (infile.exists) {
                        replaceContents2(infile);
                    }
                }
            }
        }

        var saveFolder = new Folder(targ);
        var fileName = row+1;
        var doc = activeDocument;
        var jpgOptions = new JPEGSaveOptions();  
        
        jpgOptions.quality = 8; //enter number or create a variable to set quality  
        jpgOptions.embedColorProfile = true;   
        jpgOptions.formatOptions = FormatOptions.STANDARDBASELINE;
        if(jpgOptions.formatOptions == FormatOptions.PROGRESSIVE){  
           jpgOptions.scans = 3}; //only used with Progressive  
        jpgOptions.matte = MatteType.NONE; 
       
        doc.saveAs (new File(saveFolder +'/' + fileName + '.jpg'), jpgOptions);  
        // alert("saving one image.....");
        complete++;
        percentage = complete / total * 100;
        mypercentage.text= percentage.toString().substring(0,5)+" %";
        progressWindow.update();
        app.refresh();
    } catch (ex) {
        alert("#1 " + ex + ex.line);
    }
    //finally {
    //}
    
}

function getTextExtents(text_item) {
    var ref = new ActionReference();
    ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
    var desc = executeActionGet(ref).getObjectValue(stringIDToTypeID('textKey'));
    var textS = desc.getList(stringIDToTypeID('textShape'));
    var ttt = textS.getObjectValue(0);
    var bounds = ttt.getObjectValue(stringIDToTypeID('bounds'))

    var ttt = bounds.getUnitDoubleValue(stringIDToTypeID('top'));
    var lll = bounds.getUnitDoubleValue(stringIDToTypeID('left'));
    var width = bounds.getUnitDoubleValue(stringIDToTypeID('right'));
    var height = bounds.getUnitDoubleValue(stringIDToTypeID('bottom'));
    var x_scale = 1;
    var y_scale = 1;
    if (desc.hasKey(stringIDToTypeID('transform'))) {
        var transform = desc.getObjectValue(stringIDToTypeID('transform'));
        x_scale = transform.getUnitDoubleValue(stringIDToTypeID('xx'));

        y_scale = transform.getUnitDoubleValue(stringIDToTypeID('yy'));
    }
    return {
        x: Math.round(text_item.position[0]),
        y: Math.round(text_item.position[1]),
        width: Math.round(width * x_scale),
        height: Math.round((height - ttt) * y_scale)
    };
}


////// replace contents //////
function replaceContents2(newFile) {
    //app.activeDocument.activeLayer = theLayer;
    var idplacedLayerEditContents = stringIDToTypeID("placedLayerEditContents");
    var desc1275 = new ActionDescriptor();
    executeAction(idplacedLayerEditContents, desc1275, DialogModes.NO);

    // smartdoc
    var smartdoc = app.activeDocument;
    var wid = smartdoc.width;
    var hei = smartdoc.height;

    var adimg = app.open(newFile);
    var lay = adimg.activeLayer;
    lay = lay.duplicate(smartdoc);
    adimg.close(SaveOptions.DONOTSAVECHANGES);

    var ww = lay.bounds[2] - lay.bounds[0];
    var hh = lay.bounds[3] - lay.bounds[1];
    var sclx = wid / ww;
    var scly = hei / hh;
    var scl = sclx > scly ? sclx : scly;
    lay.translate(-lay.bounds[0] + (wid - ww) * 0.5, -lay.bounds[1] + (hei - hh) * 0.5);
    lay.resize(scl * 100, scl * 100, AnchorPosition.MIDDLECENTER);

    var ff = app.activeDocument.fullName;

    app.activeDocument.flatten();

    // =======================================================
    var idsave = charIDToTypeID("save");
    var desc161 = new ActionDescriptor();
    var idIn = charIDToTypeID("In  ");
    desc161.putPath(idIn, ff);
    var idsaveStage = stringIDToTypeID("saveStage");
    var idsaveStageType = stringIDToTypeID("saveStageType");
    var idsaveBegin = stringIDToTypeID("saveBegin");
    desc161.putEnumerated(idsaveStage, idsaveStageType, idsaveBegin);
    var idDocI = charIDToTypeID("DocI");
    desc161.putInteger(idDocI, 24650);
    executeAction(idsave, desc161, DialogModes.NO);

    // =======================================================
    var idsave = charIDToTypeID("save");
    var desc162 = new ActionDescriptor();
    var idIn = charIDToTypeID("In  ");
    desc162.putPath(idIn, ff);
    var idDocI = charIDToTypeID("DocI");
    desc162.putInteger(idDocI, 24650);
    var idsaveStage = stringIDToTypeID("saveStage");
    var idsaveStageType = stringIDToTypeID("saveStageType");
    var idsaveSucceeded = stringIDToTypeID("saveSucceeded");
    desc162.putEnumerated(idsaveStage, idsaveStageType, idsaveSucceeded);
    executeAction(idsave, desc162, DialogModes.NO)


    //smartdoc.save();
    smartdoc.close(SaveOptions.DONOTSAVECHANGES);
    // =======================================================
    var idupdatePlacedLayer = stringIDToTypeID("updatePlacedLayer");
    executeAction(idupdatePlacedLayer, undefined, DialogModes.NO);
};


function setBottomTextFields(docRef, rule, docReflays) {
    if (rule === null) return;
    if (rule.length < 2) return;
    // total_info += dd.dat.length - 1;

    // var head = dd[0];
    var names = rule;
    // alert("OK_not catch");
    for (var iname = 0; iname < name.length ; iname++) {
        try {
            if (names[iname] == undefined) continue;
            // var layerRef = getLayerByName(docRef, names[iname].trim2());
            if((names[iname].indexOf("Image") !=-1 )||(names[iname].indexOf("Text") !=-1 )){//test or image
                // alert(names[iname].substring(names[iname].indexOf("#")+1,names[iname].length));
            var layerRef = activeDocument.artLayers.getByName(names[iname].substring(names[iname].indexOf("#")+1,names[iname].length));
            if (layerRef.kind == LayerKind.TEXT) {
                docRef.activeLayer = layerRef;
                var ww = getTextExtents(layerRef.textItem);
                var boundHeight0 = ww.height;
                //setbound(undefined, docRef.height);
                var gr300 = boundHeight0;
                var currentFontSize = layerRef.textItem.size;
                currentFontSize = Number(currentFontSize.toString().replace(" pt", ""));
                var leamult = 0.0;
                try {
                    leamult = Number(layerRef.textItem.leading.toString().replace(" pt", ""));
                    leamult = leamult / currentFontSize;
                } catch (_) {}

                var lay = {
                    "lay": layerRef,
                    "col": iname,
                    "gr300": gr300,
                    "wwheight": ww.height,
                    "size": currentFontSize,
                    "lea": leamult
                };
                docReflays.push(lay);
            } else if (layerRef.kind == LayerKind.SMARTOBJECT) {
                var lay = {
                    "lay": layerRef,
                    "col": iname
                };
                docReflays.push(lay);
            }
        }
        } catch (e) {
            // alert("this 1 error"+e+e.line);
        }
    }
    
}


function createNamedSnapshot(name) {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putClass(charIDToTypeID('SnpS'));
    desc.putReference(charIDToTypeID('null'), ref);
    var ref1 = new ActionReference();
    ref1.putProperty(charIDToTypeID('HstS'), charIDToTypeID('CrnH'));
    desc.putReference(cTID('From'), ref1);
    desc.putString(charIDToTypeID('Nm  '), name);
    desc.putEnumerated(charIDToTypeID('Usng'), charIDToTypeID('HstS'), charIDToTypeID('FllD'));
    executeAction(charIDToTypeID('Mk  '), desc, DialogModes.NO);
};



function revertNamedSnapshot(name) {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putName(charIDToTypeID('SnpS'), name);

    desc.putReference(charIDToTypeID('null'), ref);
    executeAction(charIDToTypeID('slct'), desc, DialogModes.NO);
};

function getLayerByName(obj, name) {
    alert("find?");
    for (var i = 0; i < obj.artLayers.length; i++) {
        if (obj.artLayers[i].name.toUpperCase() == name.toUpperCase()) return obj.artLayers[i];
    }
    for (var i = 0; i < obj.layerSets.length; i++) {
        getLayerByName(obj.layerSets[i], name);
    }
    return null;
}


function fitText(gr300, currentFontSize, lea, layerRef) {
    var step = 25.0;
    var len = calcCounts();

    var count = 0;
    while (true) {
        var currentFontSize0 = currentFontSize;
        currentFontSize += step;
        fontSize(currentFontSize, lea);

        app.activeDocument.activeLayer = layerRef;
        var lll = calcCounts();
        if (lll < len) {
            step *= 0.3;
            currentFontSize = currentFontSize0;
        }
        if (step < 0.01) break;
    }
    while (calcCounts() < len) {
        currentFontSize -= step;
        fontSize(currentFontSize, lea);
    }
}



function fontSize(v, lea) {
    function sTT(v) {
        return stringIDToTypeID(v)
    }(ref1 = new ActionReference()).putProperty(sTT('property'), tS = sTT('textStyle'))
    ref1.putEnumerated(sTT('textLayer'), sTT('ordinal'), sTT('targetEnum'));
    (dsc1 = new ActionDescriptor()).putReference(sTT('null'), ref1);
    (dsc2 = new ActionDescriptor()).putInteger(sTT('textOverrideFeatureName'), 808465458), dsc2.putInteger(sTT('typeStyleOperationType'), 3)
    dsc2.putUnitDouble(sTT('size'), sTT('pixelsUnit'), v), dsc1.putObject(sTT('to'), tS, dsc2), executeAction(sTT('set'), dsc1, DialogModes.NO);

    if (lea > 0.01) {
        // =======================================================
        var idsetd = charIDToTypeID("setd");
        var desc162 = new ActionDescriptor();
        var idnull = charIDToTypeID("null");
        var ref11 = new ActionReference();
        var idPrpr = charIDToTypeID("Prpr");
        var idTxtS = charIDToTypeID("TxtS");
        ref11.putProperty(idPrpr, idTxtS);
        var idTxLr = charIDToTypeID("TxLr");
        var idOrdn = charIDToTypeID("Ordn");
        var idTrgt = charIDToTypeID("Trgt");
        ref11.putEnumerated(idTxLr, idOrdn, idTrgt);
        desc162.putReference(idnull, ref11);
        var idT = charIDToTypeID("T   ");
        var desc163 = new ActionDescriptor();
        var idtextOverrideFeatureName = stringIDToTypeID("textOverrideFeatureName");
        desc163.putInteger(idtextOverrideFeatureName, 808465461);
        var idtypeStyleOperationType = stringIDToTypeID("typeStyleOperationType");
        desc163.putInteger(idtypeStyleOperationType, 3);
        var idautoLeading = stringIDToTypeID("autoLeading");
        desc163.putBoolean(idautoLeading, false);
        var idLdng = charIDToTypeID("Ldng");
        var idPnt = charIDToTypeID("#Pnt");
        desc163.putUnitDouble(idLdng, idPnt, lea * v);
        var idTxtS = charIDToTypeID("TxtS");
        desc162.putObject(idT, idTxtS, desc163);
        executeAction(idsetd, desc162, DialogModes.NO);
    }
}


function fitText(gr300, currentFontSize, lea, layerRef) {
    var step = 25.0;
    var len = calcCounts();

    var count = 0;
    while (true) {
        var currentFontSize0 = currentFontSize;
        currentFontSize += step;
        fontSize(currentFontSize, lea);

        app.activeDocument.activeLayer = layerRef;
        var lll = calcCounts();
        if (lll < len) {
            step *= 0.3;
            currentFontSize = currentFontSize0;
        }
        if (step < 0.01) break;
    }
    while (calcCounts() < len) {
        currentFontSize -= step;
        fontSize(currentFontSize, lea);
    }
}


function makeid() {
    var d = new Date();
    var text = "";
    var possible = "0123456789";
    for (var i = 0; i < 6; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return (d.getFullYear().toString()) + ((d.getMonth() + 1).toString()) + (d.getDate().toString()) + '_' + text;
}


function calcCounts() {
    var savedState = app.activeDocument.activeHistoryState;
    //var bbb = app.activeDocument.activeLayer.textItem.fauxBold;
    //if (bbb) app.activeDocument.activeLayer.textItem.fauxBold = false;
    // =======================================================
    var idMk = charIDToTypeID("Mk  ");
    var desc335 = new ActionDescriptor();
    var idnull = charIDToTypeID("null");
    var ref53 = new ActionReference();
    var idPath = charIDToTypeID("Path");
    ref53.putClass(idPath);
    desc335.putReference(idnull, ref53);
    var idFrom = charIDToTypeID("From");
    var ref54 = new ActionReference();
    var idTxLr = charIDToTypeID("TxLr");
    var idOrdn = charIDToTypeID("Ordn");
    var idTrgt = charIDToTypeID("Trgt");
    ref54.putEnumerated(idTxLr, idOrdn, idTrgt);
    desc335.putReference(idFrom, ref54);
    executeAction(idMk, desc335, DialogModes.NO);
    var ccc = app.activeDocument.pathItems[0].subPathItems.length
    app.activeDocument.activeHistoryState = savedState;
    return ccc;
}

}();