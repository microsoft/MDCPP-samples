const Path = require('path');
const DuplicatePackageCheckerPlugin = require('duplicate-package-checker-webpack-plugin');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const { merge } = require('webpack-merge');
const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');
const dotenv = require('dotenv');

const modeConfig = env => require(`./webpack.${env}`)();

module.exports = ({ mode } = { mode: "development" }) => {
  return merge({
    entry: {
      desktop: './src/desktop.ts',
    },
    mode,
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
      fallback: {
        "crypto": false
      },
    },
    optimization: {
      splitChunks: {
        automaticNameDelimiter: '.',
      }
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          loader: "ts-loader"
        },
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              plugins: [
                '@babel/plugin-transform-optional-chaining',
                '@babel/plugin-transform-nullish-coalescing-operator'
              ]
            }
          }
        },
        {
          test: /\.scss$/,
          use: [
            'style-loader',
            'css-loader',
            'sass-loader'
          ]
        }
      ]
    },
    output: {
      filename: '[name].js',
      chunkFilename: '[name].bundle.js',
      path: Path.resolve(__dirname, 'dist'),
      libraryTarget: 'var',
      library: '[name]'
    },
    plugins: [
      new webpack.DefinePlugin({ 'process.env': JSON.stringify(dotenv.config().parsed) }),
      new BundleAnalyzerPlugin({
        analyzerMode: 'disabled',
        generateStatsFile: true,
        statsOptions: { source: false }
      }),
      new DuplicatePackageCheckerPlugin(),
      new CleanWebpackPlugin({  verbose: true, }), 
      new HtmlWebpackPlugin({
        title: 'Hello World',
        chunks: ['main']
      }),
    ]
  },
    modeConfig(mode)
  );
};