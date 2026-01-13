module.exports = function (config) {
  config.set({
    basePath: "",
    frameworks: ["mocha", "chai"],
    files: [
      { pattern: "node_modules/jquery/dist/jquery.js", watched: false },
      { pattern: "node_modules/jszip/dist/jszip.min.js", watched: false },
      { pattern: "src/jquery.table2excel.js", watched: true },
      { pattern: "test/fixtures/*.html", watched: true, included: false, served: true },
      { pattern: "test/**/*.spec.js", watched: true }
    ],
    reporters: ["progress"],
    port: 9876,
    colors: true,
    logLevel: config.LOG_INFO,
    autoWatch: false,
    browsers: ["ChromeHeadless"],
    singleRun: false,
    client: {
      mocha: {
        timeout: 5000
      }
    }
  });
};
