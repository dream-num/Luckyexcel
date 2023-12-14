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

// rollup packaging, processing es6 modules
const { rollup } = require('rollup');
// rollup typescript
const typescript = require('rollup-plugin-typescript2');
// rollup looks for node_modules module
// const { nodeResolve } = require('@rollup/plugin-node-resolve');
// rollup converts commonjs module to es6 module
// const commonjs = require('@rollup/plugin-commonjs');
// rollup code compression
// const terser = require('rollup-plugin-terser').terser;
// rollup babel plugin, support the latest ES grammar
// const babel = require('@rollup/plugin-babel').default;

const pkg = require('./package.json');


const paths = {
    pages: ['src/*.html',"assets/**"]
};

// babel config
// const babelConfig = {
//     babelHelpers: 'bundled',
//     exclude: 'node_modules/**', // Only compile our source code
//     plugins: [
//     ],
//     presets: [
//         ['@babel/preset-env', {
//             useBuiltIns: 'usage',
//             corejs: 3,
//             targets: {
//                 chrome: 58,
//                 ie: 11
//             }
//         }],
//         '@babel/preset-typescript'
//     ]
// };

// Copy html
function copyHtml(){
    return src(paths.pages)
        .pipe(dest("dist"));
}

// Refresh browser
function reloadBrowser(done) {
    reload();

    done();
}

// Monitoring static file changes
function watcher(done) {
    // watch static
    watch(paths.pages,{ delay: 500 }, series(copyHtml, reloadBrowser));
    done();
}

// 监听文件改变
const watchedBrowserify = watchify(browserify({
    basedir: '.',
    debug: true,
    entries: [
        'src/main.umd.ts'
    ],
    cache: {},
    packageCache: {},
    standalone:'LuckyExcel'
}).plugin(tsify));

// 开发模式，打包成es5，方便在浏览器里直接引用调试
function bundle() {
    return watchedBrowserify
        .transform('babelify', {
            presets: ['@babel/preset-env','@babel/preset-typescript'],
            extensions: ['.ts']
        })
        .bundle()
        .pipe(source('luckyexcel.umd.js'))
        .pipe(buffer())
        .pipe(sourcemaps.init({loadMaps: true}))
        // .pipe(uglify()) //Development environment does not need to compress code
        .pipe(sourcemaps.write('./'))
        .pipe(dest("dist"));
}

// 生产模式，打包成ES模块和Commonjs模块
async function compile() {
    
    const bundle = await rollup({
        input: 'src/main.esm.ts',
        plugins: [
            // nodeResolve(), // tells Rollup how to find date-fns in node_modules
            // commonjs(), // converts date-fns to ES modules
            typescript({
                tsconfigOverride: { 
                    compilerOptions : { module: "ESNext" } 
                }
            }),
            // terser(), // minify, but only in production
            // babel(babelConfig)
        ],
    });

    bundle.write({
        file: pkg.module,
        format: 'esm',
        name: 'LuckyExcel',
        inlineDynamicImports:true,
        // sourcemap: true
    })
    bundle.write({
        file: pkg.main,
        format: 'cjs',
        name: 'LuckyExcel',
        inlineDynamicImports:true,
        // sourcemap: true
    })
    // bundle.write({
    //     file: pkg.browser,
    //     format: 'umd',
    //     name: 'LuckyExcel',
    //     inlineDynamicImports:true,
    //     // sourcemap: true
    // })
}

// 生产模式，打包成UMD模块
function bundleUMD() {
    return browserify({
        basedir: '.',
        entries: ['src/main.umd.ts'],
        cache: {},
        packageCache: {},
        standalone:'LuckyExcel'
    })
    .plugin(tsify)
    .transform('babelify', {
        presets: ['@babel/preset-env','@babel/preset-typescript'],
        extensions: ['.ts']
    })
    .bundle()
    .pipe(source('luckyexcel.umd.js'))
    .pipe(buffer())
    // .pipe(sourcemaps.init({loadMaps: true})) //The production environment does not need source map file
    // .pipe(uglify())
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
const dev = series(clean, copyHtml, bundle, watcher, serve);

const build = series(clean, parallel(copyHtml, compile, bundleUMD));

// 每次TypeScript文件改变时Browserify会执行bundle函数
watchedBrowserify.on("update", series(bundle, reload));

// 将日志打印到控制台
watchedBrowserify.on("log", log);


exports.dev = dev;
exports.build = build;
exports.default = dev;