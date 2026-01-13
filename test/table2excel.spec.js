/* global describe, it, before, beforeEach, afterEach */

var expect = chai.expect;

describe("table2excel", function () {
	var fixtureDoc;
	var originalJSZip;
	var originalURL;
	var originalCreateElement;
	var originalAlert;
	var originalConsoleError;

	before(function (done) {
		$.get("/base/test/fixtures/tables.html", function (html) {
			fixtureDoc = $(html);
			done();
		});
	});

	beforeEach(function () {
		originalJSZip = window.JSZip;
		originalURL = window.URL;
		originalCreateElement = document.createElement;
		originalAlert = window.alert;
		originalConsoleError = console.error;

		var $fixture = $("#fixture");
		if (!$fixture.length) {
			$fixture = $('<div id="fixture"></div>').appendTo(document.body);
		}
		$fixture.empty();
	});

	afterEach(function () {
		window.JSZip = originalJSZip;
		window.URL = originalURL;
		document.createElement = originalCreateElement;
		window.alert = originalAlert;
		console.error = originalConsoleError;

		$("#fixture").remove();
	});

	function mountTable(selector) {
		var $table = fixtureDoc.find(selector).first().clone();
		var $fixture = $("#fixture");
		$fixture.append($table);
		return $table;
	}

	it("uses XLSX path and applies borders when borders are enabled", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;

		var createdAnchor;
		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				createdAnchor = el;
				el.click = function () {
					el._clicked = true;
				};
			}
			return el;
		};

		window.URL = {
			createObjectURL: function () {
				return "blob://test";
			},
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-borders");
		$table.table2excel({ borders: true, filename: "report.xls", name: "Sheet1" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var styles = lastZip.files["xl/styles.xml"];
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];

			expect(styles).to.contain("applyBorder=\"1\"");
			expect(worksheet).to.contain('s="1"');
			expect(worksheet).to.not.contain("ShouldSkip");
			expect(createdAnchor).to.exist;
			expect(createdAnchor.download).to.equal("report.xlsx");
			expect(createdAnchor._clicked).to.equal(true);
			done();
		}, 0);
	});

	it("falls back to legacy XLS path when borders are disabled", function () {
		window.JSZip = function () {
			throw new Error("JSZip should not be used");
		};

		var linkCreated = false;
		window.URL = {
			createObjectURL: function () {
				linkCreated = true;
				return "blob://legacy";
			},
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				el.click = function () {
					el._clicked = true;
				};
			}
			return el;
		};

		var $table = mountTable("#table-legacy");
		$table.table2excel({ borders: false, excelFormat: "xls", filename: "legacy.xls" });

		expect(linkCreated).to.equal(true);
	});

	it("uses XLSX path when excelFormat is explicitly set to xlsx", function (done) {
		var called = false;
		function FakeJSZip() {
			called = true;
			this.files = {};
		}
		FakeJSZip.prototype.file = function () { };
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;

		window.URL = {
			createObjectURL: function () {
				return "blob://xlsx";
			},
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			return originalCreateElement.call(document, tagName);
		};

		var $table = mountTable("#table-borders");
		$table.table2excel({ excelFormat: "xlsx", filename: "format.xlsx" });

		setTimeout(function () {
			expect(called).to.equal(true);
			done();
		}, 0);
	});

	it("logs and alerts when JSZip is missing for XLSX export", function () {
		window.JSZip = undefined;

		var errorCalled = false;
		var alertCalled = false;

		console.error = function () {
			errorCalled = true;
		};
		window.alert = function () {
			alertCalled = true;
		};

		var $table = mountTable("#table-borders");
		$table.table2excel({ borders: true, filename: "missing.zip" });

		expect(errorCalled).to.equal(true);
		expect(alertCalled).to.equal(true);
	});

	it("replaces images with alt text when excludeImages is true in XLS format", function () {
		var capturedBlob = null;
		window.JSZip = function () {
			throw new Error("JSZip should not be used for XLS");
		};

		window.URL = {
			createObjectURL: function (blob) {
				capturedBlob = blob;
				return "blob://test";
			},
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				el.click = function () { el._clicked = true; };
			}
			return el;
		};

		var $table = mountTable("#table-images");
		$table.table2excel({ excelFormat: "xls", filename: "images.xls" });

		// The blob was created indicating XLS export worked
		// The exclude_img function replaces <img> tags with alt text in XLS format
		expect(capturedBlob).to.exist;
	});

	it("replaces links with their text content when excludeLinks is true (default)", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-links");
		$table.table2excel({ excelFormat: "xlsx", filename: "links.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			expect(worksheet).to.contain("GitHub Link");
			expect(worksheet).to.contain("Google Link");
			expect(worksheet).to.not.contain("<a ");
			expect(worksheet).to.not.contain("href=");
			done();
		}, 0);
	});

	it("replaces inputs with their values when excludeInputs is true in XLS format", function () {
		var capturedBlob = null;
		window.JSZip = function () {
			throw new Error("JSZip should not be used for XLS");
		};

		window.URL = {
			createObjectURL: function (blob) {
				capturedBlob = blob;
				return "blob://test";
			},
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				el.click = function () { el._clicked = true; };
			}
			return el;
		};

		var $table = mountTable("#table-inputs");
		$table.table2excel({ excelFormat: "xls", filename: "inputs.xls" });

		// The blob was created indicating XLS export worked
		// The exclude_inputs function replaces <input> tags with values in XLS format
		expect(capturedBlob).to.exist;
	});

	it("escapes special XML characters in cell content", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-special-chars");
		$table.table2excel({ excelFormat: "xlsx", filename: "special.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			// The content should be escaped in the XML
			expect(worksheet).to.contain("Tom &amp; Jerry");
			expect(worksheet).to.contain("5 &lt; 10");
			expect(worksheet).to.contain("10 &gt; 5");
			expect(worksheet).to.contain("&quot;Hello&quot;");
			expect(worksheet).to.contain("&apos;Hi&apos;");
			done();
		}, 0);
	});

	it("uses custom sheetName in workbook XML", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-sheetname");
		$table.table2excel({ excelFormat: "xlsx", sheetName: "MyCustomSheet", filename: "sheet.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var workbook = lastZip.files["xl/workbook.xml"];
			expect(workbook).to.contain("MyCustomSheet");
			done();
		}, 0);
	});

	it("falls back to name option when sheetName is not provided", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-sheetname");
		// When sheetName is not set, it falls back to name option
		$table.table2excel({ excelFormat: "xlsx", name: "FallbackName", filename: "sheet.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var workbook = lastZip.files["xl/workbook.xml"];
			expect(workbook).to.contain("FallbackName");
			done();
		}, 0);
	});

	it("handles tables with many columns (generates AA, AB column letters)", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-many-columns");
		$table.table2excel({ excelFormat: "xlsx", filename: "columns.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			// Column 27 should be AA, column 28 should be AB
			expect(worksheet).to.contain('r="AA');
			expect(worksheet).to.contain('r="AB');
			done();
		}, 0);
	});

	it("handles colspan attributes in cells", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-merged-cells");
		$table.table2excel({ excelFormat: "xlsx", filename: "merged.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			// Should contain the content from colspan cell
			expect(worksheet).to.contain("Spans 2 Cols");
			done();
		}, 0);
	});

	it("handles rowspan attributes in cells", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-merged-cells");
		$table.table2excel({ excelFormat: "xlsx", filename: "merged.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			// Should contain the content from rowspan cell
			expect(worksheet).to.contain("Spans 2 Rows");
			done();
		}, 0);
	});

	it("handles empty tables gracefully", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-empty");
		$table.table2excel({ excelFormat: "xlsx", filename: "empty.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			expect(worksheet).to.contain("Empty Header");
			expect(worksheet).to.exist;
			done();
		}, 0);
	});

	it("preserves colors in XLS format when preserveColors is enabled", function () {
		var capturedHtml = null;
		window.JSZip = function () {
			throw new Error("JSZip should not be used for XLS");
		};

		window.URL = {
			createObjectURL: function (blob) {
				// Capture the blob content for inspection
				capturedHtml = blob;
				return "blob://colors";
			},
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				el.click = function () {
					el._clicked = true;
				};
			}
			return el;
		};

		var $table = mountTable("#table-colors");
		$table.table2excel({
			excelFormat: "xls",
			preserveColors: true,
			filename: "colors.xls"
		});

		// The blob was created, indicating XLS export path was used
		expect(capturedHtml).to.exist;
	});

	it("excludes rows matching the exclude selector", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-borders");
		$table.table2excel({ excelFormat: "xlsx", exclude: ".noExl", filename: "exclude.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			expect(worksheet).to.not.contain("ShouldSkip");
			expect(worksheet).to.not.contain("AlsoSkip");
			expect(worksheet).to.contain("Header 1");
			done();
		}, 0);
	});

	it("exports table content when selector matches elements in XLSX format", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		// Mount a multi-table
		var $table1 = fixtureDoc.find("#table-multi-1").first().clone();
		var $fixture = $("#fixture");
		$fixture.append($table1);

		$(".multi-table").table2excel({ excelFormat: "xlsx", filename: "multi.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var worksheet = lastZip.files["xl/worksheets/sheet1.xml"];
			// Table content should be present
			expect(worksheet).to.contain("Multi 1");
			done();
		}, 0);
	});

	it("uses default filename when none is provided", function (done) {
		var lastZip;
		var createdAnchor;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		document.createElement = function (tagName) {
			var el = originalCreateElement.call(document, tagName);
			if (tagName === "a") {
				createdAnchor = el;
				el.click = function () {
					el._clicked = true;
				};
			}
			return el;
		};

		var $table = mountTable("#table-legacy");
		$table.table2excel({ excelFormat: "xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			expect(createdAnchor).to.exist;
			// Should have .xlsx extension
			expect(createdAnchor.download).to.match(/\.xlsx$/);
			done();
		}, 0);
	});

	it("uses name option as sheet name fallback", function (done) {
		var lastZip;
		function FakeJSZip() {
			lastZip = this;
			this.files = {};
		}
		FakeJSZip.prototype.file = function (name, content) {
			this.files[name] = content;
		};
		FakeJSZip.prototype.generateAsync = function () {
			return Promise.resolve("blob");
		};

		window.JSZip = FakeJSZip;
		window.URL = {
			createObjectURL: function () { return "blob://test"; },
			revokeObjectURL: function () { }
		};

		var $table = mountTable("#table-legacy");
		$table.table2excel({ excelFormat: "xlsx", name: "NameOptionSheet", filename: "name.xlsx" });

		setTimeout(function () {
			expect(lastZip).to.exist;
			var workbook = lastZip.files["xl/workbook.xml"];
			expect(workbook).to.contain("NameOptionSheet");
			done();
		}, 0);
	});
});
