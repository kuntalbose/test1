!function () {

    var tablesToExcel = {};

    tablesToExcel.export = (function () {   //shift to common
        var uri = 'data:application/vnd.ms-excel;base64,'
            , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
                + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Axel Richter</Author><Created>{created}</Created></DocumentProperties>'
                + '<Styles>'
                + '<Style ss:ID="Currency"><NumberFormat ss:Format="Currency"></NumberFormat></Style>'
                + '<Style ss:ID="Date"><NumberFormat ss:Format="Medium Date"></NumberFormat></Style>'
                + '</Styles>'
                + '{worksheets}</Workbook>'
            , tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table>'
                + '<Column ss:Width="150"/> <Column   ss:Width="150"/> <Column  ss:Width="150"/> <Column  ss:Width="150"/> <Column ss:Width="150"/> '
                + '{rows}</Table></Worksheet>'
            , tmplCellXML = '<Cell{attributeStyleID}{attributeFormula}><Data ss:Type="{nameType}">{data}</Data></Cell>'
            , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
            , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        return function (tables, wsnames, wbname, appname) {
            var ctx = "";
            var workbookXML = "";
            var worksheetsXML = "";
            var rowsXML = "";

            for (var i = 0; i < tables.length; i++) {
                if (!tables[i].nodeType) tables[i] = document.getElementById(tables[i]);
                for (var j = 0; j < tables[i].rows.length; j++) {

                    rowsXML += '<Row>'
                    for (var k = 0; k < tables[i].rows[j].cells.length; k++) {
                        var dataType = tables[i].rows[j].cells[k].getAttribute("data-type");
                        var dataStyle = tables[i].rows[j].cells[k].getAttribute("data-style");
                        var dataValue = tables[i].rows[j].cells[k].getAttribute("data-value");
                        dataValue = (dataValue) ? dataValue : tables[i].rows[j].cells[k].innerHTML;
                        var dataFormula = tables[i].rows[j].cells[k].getAttribute("data-formula");
                        dataFormula = (dataFormula) ? dataFormula : (appname == 'Calc' && dataType == 'DateTime') ? dataValue : null;
                        //console.log(dataType, dataStyle, dataValue, dataFormula)

                        ctx = {
                            attributeStyleID: (dataStyle == 'Currency' || dataStyle == 'Date') ? ' ss:StyleID="' + dataStyle + '"' : ''
                            , nameType: (dataType == 'Number' || dataType == 'DateTime' || dataType == 'Boolean' || dataType == 'Error') ? dataType : 'String'
                            , data: (dataFormula) ? '' : dataValue
                            , attributeFormula: (dataFormula) ? ' ss:Formula="' + dataFormula + '"' : ''
                        };

                        rowsXML += format(tmplCellXML, ctx);
                    }
                    rowsXML += '</Row>'
                }

                ctx = { rows: rowsXML, nameWS: wsnames[i] || 'Sheet' + i };
                worksheetsXML += format(tmplWorksheetXML, ctx);
                rowsXML = "";
            }

            ctx = { created: (new Date()).getTime(), worksheets: worksheetsXML };
            workbookXML = format(tmplWorkbookXML, ctx);

            var link = document.createElement("A");
            link.href = uri + base64(workbookXML);
            link.download = wbname || 'Workbook.xls';
            link.target = '_blank';

            if (detectIE() != "ie" && detectIE() != "edge") {    // ie 11 & below

                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            } else {
                try {

                    var canvas = document.createElement('canvas');
                    canvas.id = 'someId';
                    document.body.appendChild(canvas);

                    var blob = new Blob([workbookXML], { type: 'application/vnd.ms-excel' });

                    navigator.msSaveBlob(blob, link.download);
                    document.body.removeChild(canvas);
                }
                catch (err) {
                    console.log(err.message);
                }
                //
            }

        }


    })();


    tablesToExcel.exportJson = (function () {   //shift to common
        var uri = 'data:application/vnd.ms-excel;base64,'
            , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
                + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Axel Richter</Author><Created>{created}</Created></DocumentProperties>'
                + '<Styles>'
                + '<Style ss:ID="Currency"><NumberFormat ss:Format="Currency"></NumberFormat></Style>'
                + '<Style ss:ID="Date"><NumberFormat ss:Format="Medium Date"></NumberFormat></Style>'
                + '</Styles>'
                + '{worksheets}</Workbook>'
            , tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table>'
                + '<Column ss:Width="150"/> <Column   ss:Width="150"/> <Column  ss:Width="150"/> <Column  ss:Width="150"/> <Column ss:Width="150"/> '
                + '{rows}</Table></Worksheet>'
            , tmplCellXML = '<Cell{attributeStyleID}{attributeFormula}><Data ss:Type="{nameType}">{data}</Data></Cell>'
            , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
            , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        return function (tables, wsnames, wbname, appname) {
            var ctx = "";
            var workbookXML = "";
            var worksheetsXML = "";
            var rowsXML = "";

            for (var i = 0; i < tables.length; i++) {
               
                tables[i].forEach(function (tabledata, index) {
                    if (index == 0) {
                        rowsXML += '<Row>'
                        for (var property in tabledata) {
                            var dataType = null;
                            var dataStyle = null;
                            var dataValue = property;
                            //        dataValue = (dataValue) ? dataValue : tables[i].rows[j].cells[k].innerHTML;
                            var dataFormula = null;
                            //        dataFormula = (dataFormula) ? dataFormula : (appname == 'Calc' && dataType == 'DateTime') ? dataValue : null;

                            ctx = {
                                attributeStyleID: (dataStyle == 'Currency' || dataStyle == 'Date') ? ' ss:StyleID="' + dataStyle + '"' : ''
                                , nameType: (dataType == 'Number' || dataType == 'DateTime' || dataType == 'Boolean' || dataType == 'Error') ? dataType : 'String'
                                , data: (dataFormula) ? '' : dataValue
                                , attributeFormula: (dataFormula) ? ' ss:Formula="' + dataFormula + '"' : ''
                            };

                            rowsXML += format(tmplCellXML, ctx);
                        }
                        rowsXML += '</Row>'
                    }
                })

                tables[i].forEach(function (tabledata, index) {
                   
                    rowsXML += '<Row>'
                    for (var property in tabledata) {
                        var dataType = null;
                        var dataStyle = null;
                        var dataValue = tabledata[property];
                        //        dataValue = (dataValue) ? dataValue : tables[i].rows[j].cells[k].innerHTML;
                        var dataFormula = null;
                        //        dataFormula = (dataFormula) ? dataFormula : (appname == 'Calc' && dataType == 'DateTime') ? dataValue : null;

                        ctx = {
                            attributeStyleID: (dataStyle == 'Currency' || dataStyle == 'Date') ? ' ss:StyleID="' + dataStyle + '"' : ''
                            , nameType: (dataType == 'Number' || dataType == 'DateTime' || dataType == 'Boolean' || dataType == 'Error') ? dataType : 'String'
                            , data: (dataFormula) ? '' : dataValue
                            , attributeFormula: (dataFormula) ? ' ss:Formula="' + dataFormula + '"' : ''
                        };

                                rowsXML += format(tmplCellXML, ctx);
                    }
                    rowsXML += '</Row>'
                })




                ctx = { rows: rowsXML, nameWS: wsnames[i] || 'Sheet' + i };
                worksheetsXML += format(tmplWorksheetXML, ctx);
                rowsXML = "";
            }

            ctx = { created: (new Date()).getTime(), worksheets: worksheetsXML };
            workbookXML = format(tmplWorkbookXML, ctx);

            var link = document.createElement("A");
            link.href = uri + base64(workbookXML);
            link.download = wbname || 'Workbook.xls';
            link.target = '_blank';

            if (detectIE() != "ie" && detectIE() != "edge") {    // ie 11 & below

                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            } else {
                try {

                    var canvas = document.createElement('canvas');
                    canvas.id = 'someId';
                    document.body.appendChild(canvas);

                    var blob = new Blob([workbookXML], { type: 'application/vnd.ms-excel' });

                    navigator.msSaveBlob(blob, link.download);
                    document.body.removeChild(canvas);
                }
                catch (err) {
                    console.log(err.message);
                }
                //
            }

        }


    })();

    function detectIE() {

        var ua = window.navigator.userAgent;
        
        var msie = ua.indexOf('MSIE ');
        var trident = ua.indexOf('Trident/');
        var edge = ua.indexOf('Edge/');
        var BrowserType = "";
        if (msie > 0) {
            BrowserType = "ie";  // IE 10
        }
        else if (trident > 0) {
            BrowserType = "ie";   // IE 11
        }
        else if (edge > 0) {
            BrowserType = "edge";  // edge
        }
        else {
            BrowserType = "other"; // Other browsers Chrome , Firefox  , Safari
        }

        return BrowserType;
    }


    this.tablesToExcel = tablesToExcel;
}();