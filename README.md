# jQuery table2excel Plugin by Rainabba (https://github.com/rainabba/jquery-table2excel)

#Credit for the core table export code concept goes to insin (met on Freenode in #javascript) and core code inspired from https://gist.github.com/insin/1031969


## Usage

1. Include jQuery:

	```html
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/2.0.0/jquery.min.js"></script>
	```

2. Include table2excel plugin's code:

	```html
	<script src="dist/jquery.table2excel.min.js"></script>
	```

3. Call the plugin:

	```javascript
	$("#yourHtmTable").table2excel({
	    exclude: ".excludeThisClass",
	    name: "Worksheet Name",
	    filename: "SomeFile" //do not include extension
	});
	```

#### [demo/](https://github.com/rainabba/jquery-table2excel/tree/master/demo)

Contains a simple HTML file to demonstrate your plugin.

#### [dist/](https://github.com/rainabba/jquery-table2excel/tree/master/dist)

This is where the generated files are stored once Grunt runs.

#### [.editorconfig](https://github.com/rainabba/jquery-table2excel/tree/master/.editorconfig)

This file is for unifying the coding style for different editors and IDEs.

> Check [editorconfig.org](http://editorconfig.org) if you haven't heard about this project yet.

#### [.jshintrc](https://github.com/rainabba/jquery-table2excel/tree/master/.jshintrc)

List of rules used by JSHint to detect errors and potential problems in JavaScript.

> Check [jshint.com](http://jshint.com/about/) if you haven't heard about this project yet.

#### [.travis.yml](https://github.com/rainabba/jquery-table2excel/tree/master/.travis.yml)

Definitions for continous integration using Travis.

> Check [travis-ci.org](http://about.travis-ci.org/) if you haven't heard about this project yet.

#### [table2excel.jquery.json](https://github.com/rainabba/jquery-table2excel/tree/master/table2excel.jquery.json)

Package manifest file used to publish plugins in jQuery Plugin Registry.

> Check this [Package Manifest Guide](http://plugins.jquery.com/docs/package-manifest/) for more details.

#### [Gruntfile.js](https://github.com/rainabba/jquery-table2excel/tree/master/Gruntfile.js)

Contains all automated tasks using Grunt.

> Check [gruntjs.com](http://gruntjs.com) if you haven't heard about this project yet.

#### [package.json](https://github.com/rainabba/jquery-table2excel/tree/master/package.json)

Specify all dependencies loaded via Node.JS.

> Check [NPM](https://npmjs.org/doc/json.html) for more details.

## Contributing

Check [CONTRIBUTING.md](https://github.com/rainabba/jquery-table2excel/blob/master/CONTRIBUTING.md)

## History

Check [Release](https://github.com/rainabba/jquery-table2excel/releases) list.

## License

[MIT License](http://zenorocha.mit-license.org/)
