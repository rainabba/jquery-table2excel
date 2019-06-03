# jQuery table2excel Plugin (https://github.com/rainabba/jquery-table2excel)

#Credit for the core table export code concept goes to insin (met on Freenode in #javascript) and core code inspired from https://gist.github.com/insin/1031969

# FIRST!!

Thanks for your interest. I haven't been able to maintain this and found the following project which looks well ahead of this one, so you may want to consider it first: [TableExport](https://github.com/clarketm/TableExport)


# DISCLAIMER

This plugin is a hack on a hack. The .xls extension is the only way [some versions] of excel will even open it, and you will get a warning about the contents which can be ignored. The plugin was developed against Chrome and other have contributed code that should allow it to work in Firefox and Safari, but inconsistently since it's a hack that's not well supported anywhere but Chrome. I would not use this in public production personally and it was developed for an Intranet application where users are on Chrome and had known versions of Excel installed and the users were educated about the warning. These users also save-as in Excel so that when the files are distributed, the end-users don't get the warning message.

## Install - Bower

Install `bower` globally
```sh
npm install -g bower
```

Install jquery-table2excel and dependencies
```
bower install jquery-table2excel --save
```

Include jquery and table2excel in your page
```html
<script src="bower_components\jquery\dist\jquery.min.js"></script>
<script src="bower_components\jquery-table2excel\dist\jquery.table2excel.min.js"></script>
```


## Install - Legacy

Include jQuery and table2excel plugin:
```html
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script>
<script src="//cdn.rawgit.com/rainabba/jquery-table2excel/1.1.0/dist/jquery.table2excel.min.js"></script>
```


## Using the plugin
```javascript
$("#yourHtmTable").table2excel({
    exclude: ".excludeThisClass",
    name: "Worksheet Name",
    filename: "SomeFile.xls", // do include extension
    preserveColors: false // set to true if you want background colors and font colors preserved
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
