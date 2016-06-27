var initialRow = 0;
var endRow = 0;
var initialColumn = 0;
var endColumn = 0;
var clipboard = {};
var currentColorElement = "";
var isMenuRow = false;
var isMenuColumn = false;
var menuIndex = 0;
var totalColumns = 10;

    var tablesToExcel = (function () {
        var uri = 'data:application/vnd.ms-excel;base64,'
        , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
          + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Axel Richter</Author><Created>{created}</Created></DocumentProperties>'
          + '<Styles>'
		  + '<Style ss:ID="font-ff0000"><Font ss:Color="#ff0000"></Font><Alignment ss:Vertical="Bottom" ss:WrapText="1"/></Style>'
		  + '<Style ss:ID="font-0B3861"><Font ss:Color="#0B3861"></Font></Style>'
		  + '<Style ss:ID="font-000000"><Font ss:Color="#000000"></Font></Style>'
		  + '<Style ss:ID="font-ffffff"><Font ss:Color="#ffffff"></Font></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_bold"><Font ss:Color="#ff0000" ss:Bold="1"></Font></Style>'
		  + '<Style ss:ID="font-0B3861_bold"><Font ss:Color="#0B3861" ss:Bold="1"></Font></Style>'
		  + '<Style ss:ID="font-000000_bold"><Font ss:Color="#000000" ss:Bold="1"></Font></Style>'
		  + '<Style ss:ID="font-ffffff_bold"><Font ss:Color="#ffffff" ss:Bold="1"></Font></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_bold_centerit"><Font ss:Color="#ff0000" ss:Bold="1"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_bold_centerit"><Font ss:Color="#0B3861" ss:Bold="1"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_bold_centerit"><Font ss:Color="#000000" ss:Bold="1"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_bold_centerit"><Font ss:Color="#ffffff" ss:Bold="1"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_centerit"><Font ss:Color="#ff0000"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_centerit"><Font ss:Color="#0B3861"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_centerit"><Font ss:Color="#000000"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_centerit"><Font ss:Color="#ffffff"></Font><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_color-0B3861"><Font ss:Color="#ff0000"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-000000"><Font ss:Color="#ff0000"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ffffff"><Font ss:Color="#ff0000"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ff0000"><Font ss:Color="#ff0000"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_color-0B3861_bold"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-000000_bold"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ffffff_bold"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ff0000_bold"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ff0000_color-0B3861_bold_centerit"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-000000_bold_centerit"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ffffff_bold_centerit"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ff0000_bold_centerit"><Font ss:Color="#ff0000" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  + '<Style ss:ID="font-ff0000_color-0B3861_centerit"><Font ss:Color="#ff0000"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-000000_centerit"><Font ss:Color="#ff0000"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ffffff_centerit"><Font ss:Color="#ff0000"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ff0000_color-ff0000_centerit"><Font ss:Color="#ff0000"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-0B3861_color-ff0000"><Font ss:Color="#0B3861"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-000000"><Font ss:Color="#0B3861"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-ffffff"><Font ss:Color="#0B3861"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-0B3861"><Font ss:Color="#0B3861"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-0B3861_color-ff0000_bold"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-000000_bold"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-ffffff_bold"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-0B3861_bold"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-0B3861_color-ff0000_bold_centerit"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-000000_bold_centerit"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-ffffff_bold_centerit"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-0B3861_bold_centerit"><Font ss:Color="#0B3861" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-0B3861_color-ff0000_centerit"><Font ss:Color="#0B3861"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-000000_centerit"><Font ss:Color="#0B3861"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-ffffff_centerit"><Font ss:Color="#0B3861"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-0B3861_color-0B3861_centerit"><Font ss:Color="#0B3861"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-000000_color-ff0000"><Font ss:Color="#000000"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-0B3861"><Font ss:Color="#000000"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-ffffff"><Font ss:Color="#000000"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-000000"><Font ss:Color="#000000"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-000000_color-ff0000_bold"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-0B3861_bold"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-ffffff_bold"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-000000_color-000000_bold"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-000000_color-ff0000_bold_centerit"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-0B3861_bold_centerit"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-ffffff_bold_centerit"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-000000_bold_centerit"><Font ss:Color="#000000" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-000000_color-ff0000_centerit"><Font ss:Color="#000000"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-0B3861_centerit"><Font ss:Color="#000000"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-ffffff_centerit"><Font ss:Color="#000000"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-000000_color-000000_centerit"><Font ss:Color="#000000"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ffffff_color-ff0000"><Font ss:Color="#ffffff"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-0B3861"><Font ss:Color="#ffffff"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-ffffff"><Font ss:Color="#ffffff"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-000000"><Font ss:Color="#ffffff"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ffffff_color-ff0000_bold"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-0B3861_bold"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-ffffff_bold"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-000000_bold"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ffffff_color-ff0000_bold_centerit"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-0B3861_bold_centerit"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-ffffff_bold_centerit"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-000000_bold_centerit"><Font ss:Color="#ffffff" ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="font-ffffff_color-ff0000_centerit"><Font ss:Color="#ffffff"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-0B3861_centerit"><Font ss:Color="#ffffff"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-ffffff_centerit"><Font ss:Color="#ffffff"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="font-ffffff_color-000000_centerit"><Font ss:Color="#ffffff"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="_color-0B3861"><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-ff0000"><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-000000"><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-ffffff"><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="_color-0B3861_bold"><Font ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-ff0000_bold"><Font ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-000000_bold"><Font ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/></Style>'
		  + '<Style ss:ID="_color-ffffff_bold"><Font ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/></Style>'
		  
		  
		  + '<Style ss:ID="_color-0B3861_bold_centerit"><Font ss:Bold="1"/><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-ff0000_bold_centerit"><Font ss:Bold="1"/><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-000000_bold_centerit"><Font ss:Bold="1"/><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-ffffff_bold_centerit"><Font ss:Bold="1"/><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="_color-0B3861_centerit"><Interior ss:Color="#0B3861" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-ff0000_centerit"><Interior ss:Color="#ff0000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-000000_centerit"><Interior ss:Color="#000000" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  + '<Style ss:ID="_color-ffffff_centerit"><Interior ss:Color="#ffffff" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
		  + '<Style ss:ID="_bold"><Font ss:Bold="1"/></Style>'
		  + '<Style ss:ID="_centerit"><Alignment ss:Horizontal="Center"/></Style>'
		  
		  
          + '</Styles>'
          + '{worksheets}</Workbook>'
        , tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table><ss:Column ss:Width="80"/>{rows}</Table></Worksheet>'
        , tmplCellXML = '<Cell{attributeStyleID}{attributeFormula}><Data ss:Type="{nameType}">{data}</Data></Cell>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        return function (wsnames, wbname, appname) {
            var ctx = "";
            var workbookXML = "";
            var worksheetsXML = "";
            var rowsXML = "";
            var tables = $('table');
            for (var i = 0; i < tables.length; i++) {
                for (var j = 1; j < tables[i].rows.length; j++) {
                    rowsXML += '<Row>'
                    for (var k = 1; k < tables[i].rows[j].cells.length; k++) {
                        var dataType = tables[i].rows[j].cells[k].getAttribute("data-type");
                        var dataStyle = tables[i].rows[j].cells[k].getAttribute("data-style");
                        var dataValue = tables[i].rows[j].cells[k].getAttribute("data-value");
						var cellColor = tables[i].rows[j].cells[k].className;
						var fontColor = cellColor.replace("content",'').replace("ui-selectee",'').replace("ui-selected",'').replace("top",'').replace("left",'').replace("first",'').replace(/\bcolor-\S+/g, '').replace("bold","").replace("centerit","").trim();
						var backgroundColor = cellColor.replace("content",'').replace("ui-selectee",'').replace("ui-selected",'').replace("top",'').replace("left",'').replace("first",'').replace(/\bfont-\S+/g, '').replace("bold","").replace("centerit","").trim();
						var bold = cellColor.replace("content",'').replace("ui-selectee",'').replace("ui-selected",'').replace("top",'').replace("left",'').replace("first",'').replace(/\bfont-\S+/g, '').replace(/\bcolor-\S+/g, '').replace("centerit","").trim();
						var center = cellColor.replace("content",'').replace("ui-selectee",'').replace("ui-selected",'').replace("top",'').replace("left",'').replace("first",'').replace(/\bfont-\S+/g, '').replace(/\bcolor-\S+/g, '').replace("bold","").trim();
						cellColor = (fontColor) ? fontColor : '';
						cellColor += (backgroundColor) ? '_' + backgroundColor : '';
						cellColor += (bold) ? '_' + bold : '';
						cellColor += (center) ? '_' + center : '';
                        dataValue = (dataValue) ? dataValue : tables[i].rows[j].cells[k].innerHTML;
                        var dataFormula = tables[i].rows[j].cells[k].getAttribute("data-formula");
                        dataFormula = (dataFormula) ? dataFormula : (appname == 'Calc' && dataType == 'DateTime') ? dataValue : null;
                        ctx = {
                            attributeStyleID: (cellColor) ? ' ss:StyleID="' + cellColor + '"' : ''
                               , nameType: (dataType == 'Number' || dataType == 'DateTime' || dataType == 'Boolean' || dataType == 'Error') ? dataType : 'String'
                               , data: (dataFormula) ? '' : dataValue.replace('<br>', '')
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

            console.log(workbookXML);

            var link = document.createElement("A");
            link.href = uri + base64(workbookXML);
            link.download = wbname || 'Workbook.xls';
            link.target = '_blank';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    })();



$(function() {
	$("table#IA").selectable();
	autosize($('textarea'));
	$("table#IA").colResizable({
		fixed:false,
		liveDrag:true,
		gripInnerHtml:"<div class='grip2'></div>", 
		draggingClass:"dragging" 
	});
	
	$("#exportit").click(function() {
		tablesToExcel(['Site Structure and Security'], (new Date().toISOString().slice(0, 10) + '_IA' + '_' + '.xls'), 'Excel');
		/*$("#exportExcel").battatech_excelexport({
			containerid: "exportExcel"
		});*/
	});
		
});


$("table#IA").on('contextmenu','.nameBar', function () {
	isMenuRow = false;
	isMenuColumn = false;
	menuIndex = 0;
	var currentRow = $(this).closest("tr")[0].rowIndex;
	if(currentRow > 0){
		isMenuRow = true;
		menuIndex = currentRow;
	}else{
		isMenuColumn = true;
		menuIndex = $(this).index();
	}
	$("#contextmenu").css({"top": event.pageY +  "px", "left": event.pageX +  "px"}).show();
	return false;
});

$(document).bind("click", function(){
	$("#contextmenu").hide();
	$("#colorMenu").hide();
});
  
$("table#IA").selectable({
  filter:'tbody td',
  selected: function(event, ui){
	var currentRow = $(ui.selected).parent().find("th").attr('data-name');
	var currentColumn = $(ui.selected).index();
	if(initialRow == 0){
		initialRow = currentRow;
		initialColumn = currentColumn;
	}
	if(currentRow > endRow)
		endRow = currentRow;
	if(currentColumn > endColumn)
		endColumn = currentColumn;
	
	$(ui.selected).parent().find("th").addClass('active');
	$("table#IA").find('tr:nth-child(1)').find('th:nth-child('+ (parseInt(currentColumn)+1) +')').addClass('active');
	
	if($(ui.selected).hasClass("bold"))
		$("#boldit").addClass("active");
	else
		$("#boldit").removeClass("active");
		
	if($(ui.selected).hasClass("leftit"))
		$("#leftit").addClass("active");
	else
		$("#leftit").removeClass("active");
		
	if($(ui.selected).hasClass("centerit"))
		$("#centerit").addClass("active");
	else
		$("#centerit").removeClass("active");
	
  },
});

$("table#IA").on('keyup','textarea', function () {
  $(this).parent().parent().find('.contentText').text($(this).val());
})

$("table#IA").selectable({
  start: function( event, ui ) {
	$('.contentText').show();
	$('textarea').hide();
	$(".ExcelCursorOuter").hide();
	initialRow = 0;
	initialColumn = 0;
	endRow = 0;
	endColumn = 0;
	$(".nameBar").removeClass("active");
  }
});

$("table#IA").selectable({
  stop: function( event, ui ) {
	$("table#IA").find('tr:nth-child(' + (parseInt(endRow)+1) + ')').find('td:nth-child('+ (parseInt(endColumn)+1) +')').find(".ExcelCursorOuter").show();
  }
});

$("table#IA").on('dblclick','.content', function () {
	if($(this).find('.contentText').text()=="")
		$(this).find('textarea').val('');
	$(this).find('.contentText').hide();
	$(this).find('textarea').show();
	$(this).find('textarea').select();
});

$(document).bind('copy', function(e) {
	var rows = (endRow - initialRow) + 1;
	var columns = (endColumn - initialColumn) + 1;
	clipboard = {};
	for (var i=0; i< rows;i++){
		for (var j=0 ; j < columns; j++){
			clipboard[j + "," + i] = {};
			clipboard[j + "," + i]["class"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').attr("class").replace(/\bcontent/g,"").replace(/\btop/g,"").replace(/\bfirst/g,"").replace(/\bui-selectee/g,"").replace(/\bui-selected/g,"");
			clipboard[j + "," + i]["content"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').find(".contentText").html();
		}
	}
}); 
$(document).bind('paste', function() {
	var positionx = "";
	var positiony = "";
	for(var cell in clipboard){
		positionx = cell.split(",")[0];
		positiony = cell.split(",")[1];
		$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + parseInt(positiony)) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + parseInt(positionx)) +')').addClass(clipboard[cell]["class"]);
		$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + parseInt(positiony)) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + parseInt(positionx)) +')').find(".contentText").html(clipboard[cell]["content"]);
		$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + parseInt(positiony)) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + parseInt(positionx)) +')').find("textarea").html(clipboard[cell]["content"]);
	}
}); 
$(document).bind('cut', function() {
	var rows = (endRow - initialRow) + 1;
	var columns = (endColumn - initialColumn) + 1;
	clipboard = {};
	for (var i=0; i< rows;i++){
		for (var j=0 ; j < columns; j++){
			clipboard[j + "," + i] = {};
			clipboard[j + "," + i]["class"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').attr("class").replace(/\bcontent/g,"").replace(/\btop/g,"").replace(/\bfirst/g,"").replace(/\bui-selectee/g,"").replace(/\bui-selected/g,"");
			$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').removeClass(clipboard[j + "," + i]["class"]);
			clipboard[j + "," + i]["content"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').find(".contentText").html();
			$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').find(".contentText").html("");
		}
	}
});

$(document).bind('contextmenu', function(e) {
	//alert('Right Click  is not allowed !!!');
	e.preventDefault();
});

$("html").keyup(function(e){
    if(e.keyCode == 46) {
        $(".ui-selected").find('.contentText').html('');
    }
});

function SelectSelectableElement (selectableContainer, elementsToSelect)
{
	$(".ui-selected").removeClass("ui-selected");
    for(var i=0;i<elementsToSelect.length;i++){
		elementsToSelect[i].addClass("ui-selecting");
	}
    selectableContainer.data("ui-selectable")._mouseStop(null);
}

$("table#IA").on('click','.nameBar.top', function () {
	var rowCount = $('table tr:last').index() + 1;
	var position = $(this).index();
	var collection = [];
	for (var i=2;i<=rowCount;i++){
		collection.push($("table#IA").find('tr:nth-child(' + (i) + ')').find('td:nth-child('+ (position+1) +')'));
	}
	SelectSelectableElement($("table#IA"),collection);
});

$("table#IA").on('click','.nameBar.left', function () {
	var ColumnCount = $(this).parent().find("> td").length + 1;
	var collection = [];
	console.log(ColumnCount);
	for (var i=2;i<=ColumnCount;i++){
		collection.push($(this).parent().find('td:nth-child('+ (i) +')'));
	}
	SelectSelectableElement($("table#IA"),collection);
});

$("#boldit").click(function() {
	if($(this).hasClass("active")){
		$(this).removeClass("active");
		$(".ui-selected").removeClass("bold");
	}else{
		$(".ui-selected").addClass("bold");
		$(this).addClass("active");
	}
});

$(".alignment").click(function() {
	if($(this).hasClass("active")){
		$(this).removeClass("active");
		$(".ui-selected").removeClass($(this).attr("id"));
	}else{
		$(".alignment").removeClass("active");
		$(".ui-selected").removeClass("leftit");
		$(".ui-selected").removeClass("centerit");
		$(".ui-selected").addClass($(this).attr("id"));
		$(this).addClass("active");
	}
});

$("#copyit").click(function() {
	var rows = (endRow - initialRow) + 1;
	var columns = (endColumn - initialColumn) + 1;
	clipboard = {};
	for (var i=0; i< rows;i++){
		for (var j=0 ; j < columns; j++){
			clipboard[j + "," + i] = {};
			clipboard[j + "," + i]["class"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').attr("class").replace(/\bcontent/g,"").replace(/\btop/g,"").replace(/\bfirst/g,"").replace(/\bui-selectee/g,"").replace(/\bui-selected/g,"");
			clipboard[j + "," + i]["content"] = $("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + i) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + j) +')').find(".contentText").html();
		}
	}
});

$("#pasteit").click(function() {
	var positionx = "";
	var positiony = "";
	for(var cell in clipboard){
		positionx = cell.split(",")[0];
		positiony = cell.split(",")[1];
		$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + parseInt(positiony)) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + parseInt(positionx)) +')').addClass(clipboard[cell]["class"]);
		$("table#IA").find('tr:nth-child(' + (parseInt(initialRow) + 1 + parseInt(positiony)) + ')').find('td:nth-child('+ (parseInt(initialColumn) + 1 + parseInt(positionx)) +')').find(".contentText").html(clipboard[cell]["content"]);
	}
});

$("#bucketSelect").click(function() {
	var color = $(this).find('#bucketcolor').attr("class");
	$(".ui-selected").attr('class', $(".ui-selected").get(0).className.replace(/\bcolor-\S+/g, ''));
	$(".ui-selected").addClass(color);
});

$("#charSelect").click(function() {
	var color = $(this).find('#charColor').attr("class");
	$(".ui-selected").attr('class', $(".ui-selected").get(0).className.replace(/\bfont-\S+/g, ''));
	$(".ui-selected").addClass(color.replace(/\bcolor-/g, 'font-'));
});

$('.selectArrow').bind("click", function(event){
	currentColorElement = $(this).parent().parent().find('div[data-element="colorize"]');
	$("#colorMenu").css({"top": event.pageY +  "px", "left": event.pageX +  "px"}).show();
	return false;
});

$(".colorDiv").click(function() {
	currentColorElement.attr("class",$(this).attr("id"));
	if(currentColorElement.attr("id")=="bucketColor"){
		$(".ui-selected").attr('class', $(".ui-selected").get(0).className.replace(/\bcolor-\S+/g, ''));
		$(".ui-selected").addClass($(this).attr("id"));
	}else{
		$(".ui-selected").attr('class', $(".ui-selected").get(0).className.replace(/\bfont-\S+/g, ''));
		$(".ui-selected").addClass($(this).attr("id").replace(/\bcolor-/g, 'font-'));
	}
});

$("#insertCol").click(function() {
	if(isMenuColumn){
		var newHeaderChar = String.fromCharCode($("table#IA").find('tr:nth-child(1)').find('th:last').attr("data-name").charCodeAt() + 1);
		$("table#IA").find('tr:nth-child(1)').find('th:last').after('<th class="nameBar top" data-name="' + newHeaderChar + '">' + newHeaderChar + '</th>');
		var currentIndex = 1;
		if(menuIndex==1){
			$("table#IA").find('tr').each(function(){
				console.log(currentIndex);
				if(currentIndex==2){
					$(this).find('td').eq(parseInt(menuIndex)-1).removeClass('first');
					$(this).find('td').eq(parseInt(menuIndex)-1).before('<td class="content first top ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>');
				}else{
					$(this).find('td').eq(parseInt(menuIndex)-1).removeClass('first');
					$(this).find('td').eq(parseInt(menuIndex)-1).before('<td class="content first ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>');
				}
				currentIndex++;
			});
		}else{
			$("table#IA").find('tr').each(function(){
				if(currentIndex==2){
					$(this).find('td').eq(parseInt(menuIndex)-1).before('<td class="content top ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>');
				}else{
					$(this).find('td').eq(parseInt(menuIndex)-1).before('<td class="content ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>');
				}
				currentIndex++;
			});
		}
		totalColumns++;
	}else{
		var newHeaderChar = parseInt($("table#IA").find('tr:nth-child(' + (menuIndex + 1) + ')').find('th').attr("data-name"));
		var rowContent = "";
		for (var i=0;i<totalColumns;i++){
			if(i==0){
				rowContent+= '<td class="content first ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>';
			}else{
				rowContent+= '<td class="content ui-selectee"><div class="contentText"></div><div class="ExcelCursorOuter"><textarea rows="1"></textarea><div class="ExcelCursor ui-state-default"></div></div></td>';
			}
		}
		$("table#IA").find('tr:nth-child(' + parseInt(menuIndex) + ')').after('<tr><th class="nameBar left" data-name="' + newHeaderChar + '">' + newHeaderChar + '</th>' + rowContent + '</tr>');
		for (var i = (menuIndex + 2); i<=$("table#IA").find('tr').length; i++){
			var newNum = parseInt($("table#IA").find('tr:nth-child(' + i + ')').find('th').attr("data-name")) + 1;
			$("table#IA").find('tr:nth-child(' + i + ')').find('th').attr("data-name",newNum);
			$("table#IA").find('tr:nth-child(' + i + ')').find('th').text(newNum);
		}
	}
	$("table#IA").colResizable({disable:true});
	$("table#IA").colResizable({
		fixed:false,
		liveDrag:true,
		gripInnerHtml:"<div class='grip2'></div>", 
		draggingClass:"dragging" 
	});
});

$("#deleteCol").click(function() {
	console.log('in');
	if(isMenuColumn){
		$("table#IA").find('tr:nth-child(1)').find('th:last').remove();
		$("table#IA").find('tr').each(function(){
				$(this).find('td').eq(parseInt(menuIndex)-1).remove();
		});
		totalColumns--;
	}else{
		for (var i = (menuIndex + 2); i<=$("table#IA").find('tr').length; i++){
			var newNum = parseInt($("table#IA").find('tr:nth-child(' + i + ')').find('th').attr("data-name")) - 1;
			$("table#IA").find('tr:nth-child(' + i + ')').find('th').attr("data-name",newNum);
			$("table#IA").find('tr:nth-child(' + i + ')').find('th').text(newNum);
		}
		$("table#IA").find('tr:nth-child(' + (menuIndex + 1) + ')').remove();
	}
});


