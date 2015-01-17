(function() {

  'use strict';

  /**
   * OOXML Tag Finder
   * Retrieves all docx, xlsx, and pptx files from a given path (non-recursive)
   * and opens each one to search for a given xml tag.
   * Reports all OOXML files which contain that tag.
   *
   * @author jelte@liebrand.co.uk (Jelte Liebrand)
   */
  var argv = require('minimist')(process.argv.slice(2), {
    alias: {
      'v': 'verbose'
    }
  });
  var _ = require('lodash');
  var Promise = require('promise');
  var AdmZip = require('adm-zip');
  var fs = require('fs');
  var xmldoc = require("xmldoc");
  var ProgressBar = require('progress');

  var searchTag = argv._[0];
  var searchPath = argv._[1];
  var MISSING_PARAM = 'Error: missing command line arguments';
  process.on('uncaughtException', logError);
  var pbar;
  var foundOfficeFileNames_ = [];


  if (!searchTag || !searchPath) {
    throw new Error(MISSING_PARAM);
  }

  // main app
  getFileList(searchPath)
      .then(filterOfficeFiles)
      .then(function(fileList) {
        pbar = new ProgressBar('Analysing: [:bar] :percent', {
          total: fileList.length,
          width: 30
        });

        fileList.forEach(function(fileName) {
          openOfficeDoc(fileName)
              .then(function(officeFile) {

                pbar.tick();

                officeFile.xmlEntries.forEach(function(xmlEntry) {

                  // now that we have an individual xml entry, do the actual
                  // search for the tag

                  var xmlData = xmlEntry.getData().toString('utf8');
                  var xml = new xmldoc.XmlDocument(xmlData);

                  var xmlPropPath = [];
                  function parseXml(el) {

                    if (el.name === searchTag) {
                      foundOfficeFileNames_.push(officeFile.fileName);

                      if (argv.verbose) {
                        console.log('%s: \tFound (in %s) \t %s',
                            officeFile.fileName,
                            xmlEntry.entryName,
                            xmlPropPath.concat([el.name]).join('.'));
                      }
                    }

                    if (el.children && el.children.length > 0) {
                      xmlPropPath.push(el.name);
                      el.children.forEach(parseXml);
                      xmlPropPath.pop();
                    }
                  }

                  parseXml(xml);

                });

              })
              .catch(logError);
        });
      })
      .then(function() {
        _.uniq(foundOfficeFileNames_).forEach(function(fileName) {
          console.log('%s', fileName);
        });
        process.exit(0);
      })
      .catch(logError);


  /**
   * Log an error to stderr. Includes logging the callstack (if there is one)
   * and prints the usage
   *
   * @param {Error} err the error object
   */
  function logError(err) {
    console.error('\n' + err);
    if (err.stack && err.message !== MISSING_PARAM) {
      console.error(err.stack);
    }
    printUsage();
    process.exit((err && err.errno) || 1);
  }


  /**
   * Print the usage for this app
   */
  function printUsage() {
    console.log('\nUsage:   parseooxml [-v] <tag> <path | file>');
    console.log('');
    console.log('Example: parseooxml w:tabs ./word/docx/\n');
  }


  /**
   * @param {String} path the path for which to retrieve a file listing
   * @return {Promise} returns a promise which resolves to a list of files in
   *     the given path
   */
  function getFileList(path) {
    process.stdout.write('Parsing filesystem...\n\n\r');

    return new Promise(function(resolve, reject) {

      walkDir(path, function(err, files) {
        if (err) {
          if (err.errno === 27) {
            // searchPath is not a path; assume its a file name
            resolve([path]);
          } else {
            reject(err);
          }
        } else {
          resolve(files);
          // resolve(files.map(function(fileName) {return path + fileName;}));
        }
      });
    });
  }


  /**
   * @param {Array} fileList list of file names (strings)
   * @return {Array} returns an filtered array of just docx/xslx/pptx file names
   */
  function filterOfficeFiles(fileList) {
    var filtered = fileList.filter(function(fileName) {
      return fileName.toLowerCase().match(/\.(?:docx|xlsx|pptx)$/);
    });
    return filtered;
  }


  /**
   * @param {String} fileName the name of an OOXML file
   * @return {Promise} returns a promise which resolves to an object containing
   *     the original file name and a reference to the Zip Structure that
   *     represents that OOXML file
   */
  function openOfficeDoc(fileName) {
    try {
      var officeFile = {
        fileName: fileName,
        zipStructure: new AdmZip(fileName)
      };
    } catch(e) {
      console.log('bingo');
      console.log(fileName);
    }
    // console.log(officeFile.zipStructure);
    // process.exit(1);

    var zipEntries = officeFile.zipStructure.getEntries();
    officeFile.xmlEntries = zipEntries.filter(function(zipEntry) {
      return zipEntry.entryName.match(/\.xml$/);
    });

    return Promise.resolve(officeFile);
  }


  /**
   * Recursively read all file names from a path
   *
   * @param {String} dir the directory to read
   * @param {Function} done the callback to use when done reading all files
   */
  function walkDir(dir, done) {
    var results = [];
    fs.readdir(dir, function(err, list) {
      if (err) {
        return done(err);
      }
      var pending = list.length;
      if (!pending) {
        return done(null, results);
      }
      list.forEach(function(file) {
        file = dir + '/' + file;
        fs.stat(file, function(err, stat) {
          if (err) {
            return done(err);
          }
          if (stat && stat.isDirectory()) {
            walkDir(file, function(err, res) {
              if (err) {
                return done(err);
              }
              results = results.concat(res);
              if (!--pending) {
                done(null, results);
              }
            });
          } else {
            results.push(file);
            if (!--pending) {
              done(null, results);
            }
          }
        });
      });
    });
  }

})();