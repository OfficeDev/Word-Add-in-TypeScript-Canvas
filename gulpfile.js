/*
Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information
*/

/* 
Gulp task configuration for transcompiling TypeScript, interpreting SASS into CSS, 
reading certs, live reload, and copy files for serving up for the add-in. The gulp-connect
server provides logging (with the debug switch) so that you can see whether calls are
being made to the add-in command function file. LiveReload is helpful as you can see
your changes reflected in a running add-in without have to restart the development
server. 
*/ 
'use strict';

var gulp = require('gulp'),
    plumber = require('gulp-plumber'),
    del = require('del'),

    sass = require('gulp-sass'),

    fs = require('fs'),
    
    typescript = require('gulp-typescript'),
    sourcemaps = require('gulp-sourcemaps'),
    tsConfigGlob = require('tsconfig-glob'),
    merge = require('merge2'),

    connect = require('gulp-connect'),
    config = require('./gulpfile.config.json'),
    packageConfig = require('./package.json'),
    errorHandler = function (error) {
        console.log(error);
        this.emit('end');
    };

// Create a TypeScript project. This is used in the compile:ts task.
var tsProject = typescript.createProject('./tsconfig.json', {
    sortdest: true
});

// Deletes all of the files in www directory. See gulpfile.config.json.
gulp.task('clean', function (done) {
    return del(config.server.root, done);
});

// Specifies glob patterns for tsconfig file.
gulp.task('ref', function () {
    return tsConfigGlob();
});

// Runs the 'clean' task and copies the packages specified in the 
// package.json overrides key to the www/lib directory. 
gulp.task('copy:libs', ['clean'], function (done) {
    return Object.keys(packageConfig.overrides)
        .map(function (key) {
            try {
                var value = packageConfig.overrides[key];
                if (Array.isArray(value)) {
                    var files = value.map(function (filePath) {
                        console.log('Copied: ', filePath);
                        return config.lib.source + "/" + key + "/" + filePath;
                    });
                    return gulp.src(files)
                        .pipe(plumber(errorHandler))
                        .pipe(gulp.dest(config.lib.dest));
                } else {
                    var file = config.lib.source + "/" + key + "/" + value;
                    console.log('Copied: ', value);
                    return gulp.src(file)
                        .pipe(plumber(errorHandler))
                        .pipe(gulp.dest(config.lib.dest));
                }
            } catch (exception) {
                console.error('Failed to load package: ', key);
            }
        });
});

// Interprets the SASS file to CSS files and puts the CSS file
// into the www directory as specified in gulpfile.config.json.
// Reloads the add-in.
gulp.task('compile:sass', function () {
    return gulp.src(config.app.source + "/**/*.scss")
        .pipe(plumber(errorHandler))
        .pipe(sass())
        .pipe(gulp.dest(config.app.dest))
        .pipe(connect.reload());
});

// Transcompiles the TypeScript files into JavaScript and places the 
// files into the www directory as specified in gulpfile.config.json.
// Reloads the add-in.
gulp.task('compile:ts', ['ref'], function () {
    var tsResult = tsProject.src()
        .pipe(sourcemaps.init())
        .pipe(plumber(errorHandler))
        .pipe(typescript(tsProject));

    tsResult.js
        .pipe(sourcemaps.write('./'))
        .pipe(gulp.dest(config.app.dest))
        .pipe(connect.reload());
});

// Copies the non TypeScript and SASS files in the src directory to the
// www directory.
gulp.task('copy:misc', function () {
    gulp.src(config.app.source + "/**/!(*.ts|*.scss)", {
            base: config.app.source
        })
        .pipe(plumber(errorHandler))
        .pipe(gulp.dest(config.app.dest));
});

// Copies files from src to www directories and reloads the add-in.
gulp.task('refresh', ['copy:misc'], function () {
    gulp.src(config.app.source + "/**/!(*.ts|*.scss)")
        .pipe(connect.reload())
        .pipe(plumber(errorHandler));
});

// Watches for changes to source files, kicks off compile tasks,
// and then reloads the add-in. This makes development much better. 
gulp.task('watch', function () {
    gulp.watch(config.app.source + "/**/*.scss", ['compile:sass']);
    gulp.watch(config.app.source + "/**/*.ts", ['compile:ts']);
    gulp.watch(config.app.source + "/**/!(*.ts|*.scss)", ['refresh']);
});

// Bundles the compile tasks.
gulp.task('compile', ['compile:sass', 'compile:ts', 'copy:misc']);

// Default gulp task. Configures the dev web server endpoint, compiles
// TypeScript and SASS, specifies certs for HTTPS. See gulpfile.config.json
// for key values.
gulp.task('default', ['compile', 'watch'], function () {
    return connect.server({
        root: config.server.root,
        host: config.server.host,
        port: config.server.port,
        https: {
            key: fs.readFileSync(config.server.serverkey),
            cert: fs.readFileSync(config.server.servercert),
            ca: fs.readFileSync(config.server.cacert),
            passphrase: config.server.passphrase
        },
        livereload: true,
        debug: true
    });
});