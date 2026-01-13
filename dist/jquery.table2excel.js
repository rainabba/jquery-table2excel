/*
 *  jQuery table2excel - v1.1.1
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *
 *  Made by rainabba
 *  Under MIT License
 */
/*
 *  jQuery table2excel - v1.1.2
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *
 *  Made by rainabba
 *  Under MIT License
 */
//table2excel.js
(function ( $, window, document, undefined ) {
    var pluginName = "table2excel",

    defaults = {
        exclude: ".noExl",
        name: "Table2Excel",
        filename: "table2excel",
        fileext: ".xls",
        exclude_img: true,
        exclude_links: true,
        exclude_inputs: true,
        preserveColors: false,
        borders: false,
        excelFormat: "xls" // "xls" or "xlsx"
    };

    // Helper functions for XLSX generation
    function generateContentTypesXML() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "</Types>";
    }

    function generateRelsXML() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>";
    }

    function generateWorkbookXML(sheetName) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets>" +
            "<sheet name=\"" + escapeXML(sheetName || "Sheet1") + "\" sheetId=\"1\" r:id=\"rId1\"/>" +
            "</sheets>" +
            "</workbook>";
    }

    function generateWorkbookRelsXML() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>";
    }

    function generateStylesXML(withBorders) {
        var borderDefinition = withBorders ?
            "<border>" +
            "<left style=\"thin\"><color rgb=\"FF000000\"/></left>" +
            "<right style=\"thin\"><color rgb=\"FF000000\"/></right>" +
            "<top style=\"thin\"><color rgb=\"FF000000\"/></top>" +
            "<bottom style=\"thin\"><color rgb=\"FF000000\"/></bottom>" +
            "</border>" :
            "<border><left/><right/><top/><bottom/></border>";

        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\">" +
            "<fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill>" +
            "</fills>" +
            "<borders count=\"2\">" +
            "<border><left/><right/><top/><bottom/></border>" +
            borderDefinition +
            "</borders>" +
            "<cellXfs count=\"2\">" +
            "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>" +
            "<xf borderId=\"" + (withBorders ? "1" : "0") + "\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"" + (withBorders ? " applyBorder=\"1\"" : "") + "/>" +
            "</cellXfs>" +
            "</styleSheet>";
    }

    function escapeXML(str) {
        if (str == null) {
            return "";
        }
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    function generateWorksheetXML(tableData, withBorders) {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>";

        var styleIndex = withBorders ? "1" : "0";

        for (var rowIndex = 0; rowIndex < tableData.length; rowIndex++) {
            var row = tableData[rowIndex];
            xml += "<row r=\"" + (rowIndex + 1) + "\">";

            for (var colIndex = 0; colIndex < row.length; colIndex++) {
                var cellRef = columnToLetter(colIndex) + (rowIndex + 1);
                var cellValue = escapeXML(row[colIndex]);

                xml += "<c r=\"" + cellRef + "\" s=\"" + styleIndex + "\" t=\"inlineStr\">" +
                    "<is><t>" + cellValue + "</t></is>" +
                    "</c>";
            }

            xml += "</row>";
        }

        xml += "</sheetData></worksheet>";
        return xml;
    }

    function columnToLetter(column) {
        var temp, letter = "";
        while (column >= 0) {
            temp = column % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            column = Math.floor(column / 26) - 1;
        }
        return letter;
    }

    // The actual plugin constructor
    function Plugin ( element, options ) {
            this.element = element;
            // jQuery has an extend method which merges the contents of two or
            // more objects, storing the result in the first object. The first object
            // is generally empty as we don't want to alter the default options for
            // future instances of the plugin
            //
            this.settings = $.extend( {}, defaults, options );
            this._defaults = defaults;
            this._name = pluginName;
            this.init();
    }

    Plugin.prototype = {
        init: function () {
            var e = this;

            // Determine if we should use XLSX format (when borders are enabled or explicitly requested)
            var useXlsx = e.settings.borders || e.settings.excelFormat === "xlsx";
            
            if (useXlsx) {
                // Extract table data as array for XLSX export
                e.extractTableDataForXlsx();
            } else {
                // Use legacy XLS export
                var utf8Heading = "<meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\">";
                e.template = {
                    head: "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\">" + utf8Heading + "<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>",
                    sheet: {
                        head: "<x:ExcelWorksheet><x:Name>",
                        tail: "</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>"
                    },
                    mid: "</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>",
                    table: {
                        head: "<table>",
                        tail: "</table>"
                    },
                    foot: "</body></html>"
                };

                e.tableRows = [];
        
                // Styling variables
                var additionalStyles = "";
                var compStyle = null;

                // get contents of table except for exclude
                $(e.element).each( function(i,o) {
                    var tempRows = "";
                    $(o).find("tr").not(e.settings.exclude).each(function (i,p) {
                        
                        // Reset for this row
                        additionalStyles = "";
                        
                        // Preserve background and text colors on the row
                        if(e.settings.preserveColors){
                            compStyle = getComputedStyle(p);
                            additionalStyles += (compStyle && compStyle.backgroundColor ? "background-color: " + compStyle.backgroundColor + ";" : "");
                            additionalStyles += (compStyle && compStyle.color ? "color: " + compStyle.color + ";" : "");
                        }

                        // Create HTML for Row
                        tempRows += "<tr style='" + additionalStyles + "'>";
                        
                        // Loop through each TH and TD
                        $(p).find("td,th").not(e.settings.exclude).each(function (i,q) { // p did not exist, I corrected

                            // Reset for this column
                            additionalStyles = "";

                            // Preserve background and text colors on the row
                            if(e.settings.preserveColors){
                                compStyle = getComputedStyle(q);
                                additionalStyles += (compStyle && compStyle.backgroundColor ? "background-color: " + compStyle.backgroundColor + ";" : "");
                                additionalStyles += (compStyle && compStyle.color ? "color: " + compStyle.color + ";" : "");
                            }

                            var rc = {
                                rows: $(this).attr("rowspan"),
                                cols: $(this).attr("colspan"),
                                flag: $(q).find(e.settings.exclude)
                            };

                            // Preserve original element type (th or td)
                            var tagName = q.tagName.toLowerCase();

                            if( rc.flag.length > 0 ) {
                                tempRows += "<td> </td>"; // exclude it!!
                            } else {
                                tempRows += "<" + tagName;
                                if( rc.rows > 0) {
                                    tempRows += " rowspan='" + rc.rows + "' ";
                                }
                                if( rc.cols > 0) {
                                    tempRows += " colspan='" + rc.cols + "' ";
                                }
                                if(additionalStyles){
                                    tempRows += " style='" + additionalStyles + "'";
                                }
                                tempRows += ">" + $(q).html() + "</" + tagName + ">";
                            }
                        });

                        tempRows += "</tr>";

                    });
                    // exclude img tags
                    if(e.settings.exclude_img) {
                        tempRows = exclude_img(tempRows);
                    }

                    // exclude link tags
                    if(e.settings.exclude_links) {
                        tempRows = exclude_links(tempRows);
                    }

                    // exclude input tags
                    if(e.settings.exclude_inputs) {
                        tempRows = exclude_inputs(tempRows);
                    }
                    e.tableRows.push(tempRows);
                });

                e.tableToExcel(e.tableRows, e.settings.name, e.settings.sheetName);
            }
        },

        extractTableDataForXlsx: function() {
            var e = this;
            var tableData = [];

            // get contents of table except for exclude
            $(e.element).each( function(i,o) {
                $(o).find("tr").not(e.settings.exclude).each(function (i,p) {
                    var rowData = [];
                    
                    // Loop through each TH and TD
                    $(p).find("td,th").not(e.settings.exclude).each(function (i,q) {
                        var rc = {
                            flag: $(q).find(e.settings.exclude)
                        };

                        if( rc.flag.length > 0 ) {
                            rowData.push(" "); // exclude it!!
                        } else {
                            var cellText = $(q).text();
                            
                            // exclude img tags
                            if(e.settings.exclude_img) {
                                cellText = cellText; // text() already excludes img
                            }
                            
                            rowData.push(cellText);
                        }
                    });

                    if (rowData.length > 0) {
                        tableData.push(rowData);
                    }
                });
            });

            e.exportToXlsx(tableData);
        },

        exportToXlsx: function(tableData) {
            var e = this;
            
            // Check if JSZip is available
            if (typeof JSZip === "undefined") {
                console.error("JSZip library is required for XLSX export. Please include JSZip before using the borders feature.");
                alert("JSZip library is required for XLSX export with borders. Please include the JSZip library on your page.");
                return;
            }

            var zip = new JSZip();
            
            // Build the OpenXML structure
            zip.file("[Content_Types].xml", generateContentTypesXML());
            zip.file("_rels/.rels", generateRelsXML());
            zip.file("xl/workbook.xml", generateWorkbookXML(e.settings.sheetName || e.settings.name));
            zip.file("xl/_rels/workbook.xml.rels", generateWorkbookRelsXML());
            zip.file("xl/styles.xml", generateStylesXML(e.settings.borders));
            zip.file("xl/worksheets/sheet1.xml", generateWorksheetXML(tableData, e.settings.borders));

            // Generate and download the file
            zip.generateAsync({ type: "blob" }).then(function(blob) {
                var filename = getFileName(e.settings);
                
                // Ensure .xlsx extension
                if (!filename.toLowerCase().endsWith(".xlsx")) {
                    filename = filename.replace(/\.(xls|xlsx)$/i, "") + ".xlsx";
                }

                var url = window.URL.createObjectURL(blob);
                var a = document.createElement("a");
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            });
        },

        tableToExcel: function (table, name, sheetName) {
            var e = this, fullTemplate="", i, link, a;

            e.format = function (s, c) {
                return s.replace(/{(\w+)}/g, function (m, p) {
                    return c[p];
                });
            };

            sheetName = typeof sheetName === "undefined" ? "Sheet" : sheetName;

            e.ctx = {
                worksheet: name || "Worksheet",
                table: table,
                sheetName: sheetName
            };

            fullTemplate= e.template.head;

            if ( $.isArray(table) ) {
                 Object.keys(table).forEach(function(i){
                      //fullTemplate += e.template.sheet.head + "{worksheet" + i + "}" + e.template.sheet.tail;
                      fullTemplate += e.template.sheet.head + sheetName + i + e.template.sheet.tail;
                });
            }

            fullTemplate += e.template.mid;

            if ( $.isArray(table) ) {
                 Object.keys(table).forEach(function(i){
                    fullTemplate += e.template.table.head + "{table" + i + "}" + e.template.table.tail;
                });
            }

            fullTemplate += e.template.foot;

            for (i in table) {
                e.ctx["table" + i] = table[i];
            }
            delete e.ctx.table;

            var isIE = navigator.appVersion.indexOf("MSIE 10") !== -1 || (navigator.userAgent.indexOf("Trident") !== -1 && navigator.userAgent.indexOf("rv:11") !== -1); // this works with IE10 and IE11 both :)
            //if (typeof msie !== "undefined" && msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // this works ONLY with IE 11!!!
            if (isIE) {
                if (typeof Blob !== "undefined") {
                    //use blobs if we can
                    fullTemplate = e.format(fullTemplate, e.ctx); // with this, works with IE
                    fullTemplate = [fullTemplate];
                    //convert to array
                    var blob1 = new Blob(fullTemplate, { type: "text/html" });
                    window.navigator.msSaveBlob(blob1, getFileName(e.settings) );
                } else {
                    //otherwise use the iframe and save
                    //requires a blank iframe on page called txtArea1
                    txtArea1.document.open("text/html", "replace");
                    txtArea1.document.write(e.format(fullTemplate, e.ctx));
                    txtArea1.document.close();
                    txtArea1.focus();
                    sa = txtArea1.document.execCommand("SaveAs", true, getFileName(e.settings) );
                }

            } else {
                var blob = new Blob([e.format(fullTemplate, e.ctx)], {type: "application/vnd.ms-excel"});
                window.URL = window.URL || window.webkitURL;
                link = window.URL.createObjectURL(blob);
                a = document.createElement("a");
                a.download = getFileName(e.settings);
                a.href = link;

                document.body.appendChild(a);

                a.click();

                document.body.removeChild(a);
            }

            return true;
        }
    };

    function getFileName(settings) {
        return ( settings.filename ? settings.filename : "table2excel" );
    }

    // Removes all img tags
    function exclude_img(string) {
        var _patt = /(\s+alt\s*=\s*"([^"]*)"|\s+alt\s*=\s*'([^']*)')/i;
        return string.replace(/<img[^>]*>/gi, function myFunction(x){
            var res = _patt.exec(x);
            if (res !== null && res.length >=2) {
                return res[2];
            } else {
                return "";
            }
        });
    }

    // Removes all link tags
    function exclude_links(string) {
        return string.replace(/<a[^>]*>|<\/a>/gi, "");
    }

    // Removes input params
    function exclude_inputs(string) {
        var _patt = /(\s+value\s*=\s*"([^"]*)"|\s+value\s*=\s*'([^']*)')/i;
        return string.replace(/<input[^>]*>|<\/input>/gi, function myFunction(x){
            var res = _patt.exec(x);
            if (res !== null && res.length >=2) {
                return res[2];
            } else {
                return "";
            }
        });
    }

    $.fn[ pluginName ] = function ( options ) {
        var e = this;
            e.each(function() {
                if ( !$.data( e, "plugin_" + pluginName ) ) {
                    $.data( e, "plugin_" + pluginName, new Plugin( this, options ) );
                }
            });

        // chain jQuery functions
        return e;
    };

})( jQuery, window, document );
