// gulp
const gulp = require("gulp");
// gulp 核心方法
const { src, dest, series, parallel, watch } = require('gulp');
// 浏览器加载Nodejs模块
const browserify = require("browserify");
// vinyl-source-stream会将Browserify的输出文件适配成gulp能够解析的格式
const source = require('vinyl-source-stream');
// Watchify启动Gulp并保持运行状态，当你保存文件时自动编译。 帮你进入到编辑-保存-刷新浏览器的循环中
const watchify = require("watchify");
// tsify是Browserify的一个插件，就像gulp-typescript一样，它能够访问TypeScript编译器
const tsify = require("tsify");
// 代码压缩混淆
const uglify = require('gulp-uglify');
// 支持sourcemaps
const sourcemaps = require('gulp-sourcemaps');
// 支持sourcemaps
const buffer = require('vinyl-buffer');
// 控制台打印日志
const log = require("fancy-log");
// 删除文件
const del = require('delete');
// 实时刷新浏览器
const browserSync = require('browser-sync').create();
const reload = browserSync.reload;

const paths = {
    pages: ['src/*.html',"assets/**"]
};

// 复制html
function copyHtml(){
    return src(paths.pages)
        .pipe(dest("dist"));
}

// 监听文件改变
const watchedBrowserify = watchify(browserify({
    basedir: '.',
    debug: true,
    entries: [
        'src/main.ts'
    ],
    cache: {},
    packageCache: {}
}).plugin(tsify));

// 打包成js
function bundle() {
    return watchedBrowserify
        .transform('babelify', {
            presets: ['@babel/preset-env','@babel/preset-typescript'],
            extensions: ['.ts']
        })
        .bundle()
        .pipe(source('luckyexcel.js'))
        .pipe(buffer())
        .pipe(sourcemaps.init({loadMaps: true}))
        .pipe(uglify())
        .pipe(sourcemaps.write('./'))
        .pipe(dest("dist"));
}

// 清除dist目录
function clean() {
    return del(['dist']);
}

// 静态服务器
function serve() {
    browserSync.init({
        server: {
            baseDir: "dist"
        }
    });
}

// 顺序执行
const dev = series(clean, copyHtml, bundle, serve);

const build = series(clean, copyHtml, bundle);

// 每次TypeScript文件改变时Browserify会执行bundle函数
watchedBrowserify.on("update", series(bundle, reload));

// 将日志打印到控制台
watchedBrowserify.on("log", log);


exports.dev = dev;
exports.build = build;
exports.default = dev;