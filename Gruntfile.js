module.exports = function(grunt) {

	
	
	grunt.initConfig({

		// Import package manifest
		pkg: grunt.file.readJSON("table2excel.jquery.json"),

		// Banner definitions
		meta: {
			banner: "/*\n" +
				" *  <%= pkg.title || pkg.name %> - v<%= pkg.version %>\n" +
				" *  <%= pkg.description %>\n" +
				" *  <%= pkg.homepage %>\n" +
				" *\n" +
				" *  Made by <%= pkg.author.name %>\n" +
				" *  Under <%= pkg.licenses[0].type %> License\n" +
				" */\n"
		},

		// Concat definitions
		concat: {
			dist: {
				src: ["src/jquery.table2excel.js"],
				dest: "dist/jquery.table2excel.js"
			},
			options: {
				banner: "<%= meta.banner %>"
			}
		},

		// Lint definitions
		jshint: {
			files: ["src/jquery.table2excel.js"],
			options: {
				jshintrc: ".jshintrc"
			}
		},

		// Minify definitions
		uglify: {
			my_target: {
				src: ["dist/jquery.table2excel.js"],
				dest: "dist/jquery.table2excel.min.js"
			},
			options: {
				banner: "<%= meta.banner %>"
			}
		},

	});
	
	require('load-grunt-tasks')(grunt);
	grunt.registerTask("default", ["jshint", "concat", "uglify"]);

};
