/*
 *  jQuery table2excel - v1.0.0
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *
 *  Made by rainabba
 *  Under MIT License
 */
//table2excel.js
;(function ( $, window, document, undefined ) {
		var pluginName = "table2excel",
				defaults = {
				exclude: ".noExl",
                name: "Table2Excel"
		};

		// The actual plugin constructor
		function Plugin ( element, options ) {
				this.element = element;
				// jQuery has an extend method which merges the contents of two or
				// more objects, storing the result in the first object. The first object
				// is generally empty as we don't want to alter the default options for
				// future instances of the plugin
				this.settings = $.extend( {}, defaults, options );
				this._defaults = defaults;
				this._name = pluginName;
				this.init();
		}

		Plugin.prototype = {
			init: function () {
				this.template = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>";
				this.tableRows = "";

				// get contents of table except for exclude
				$(this.element).find("tr").not(this.settings.exclude).each(function (i,o) {
					this.tableRows += "<tr>" + $(o).html() + "</tr>";
				});
				this.tableToExcel(this.tableRows, this.settings.name);
			},
			tableToExcel: function (table, name) {
				this.uri = "data:application/vnd.ms-excel;base64,";
				this.base64 = function (s) {
					return window.btoa(unescape(encodeURIComponent(s)));
				};
				this.format = function (s, c) {
					return s.replace(/{(\w+)}/g, function (m, p) {
						return c[p];
					});
				};
				this.ctx = {
					worksheet: name || "Worksheet",
					table: table
				};
				window.location.href = this.uri + this.base64(this.format(this.template, this.ctx));
			}
		};

		$.fn[ pluginName ] = function ( options ) {
				this.each(function() {
						if ( !$.data( this, "plugin_" + pluginName ) ) {
								$.data( this, "plugin_" + pluginName, new Plugin( this, options ) );
						}
				});

				// chain jQuery functions
				return this;
		};

})( jQuery, window, document );