const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require('webpack');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react'
      ],
      taskpane: [
        'react-hot-loader/patch',
        './src/taskpane/taskpane.ts',
      ],
      onsend: './src/onsend/onsend.ts',
      dialog: './src/onsend/dialog.tsx'
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      fallback: {
        "buffer": require.resolve("buffer")
      },
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          use: [
            'react-hot-loader/webpack',
            'ts-loader'
          ],
          exclude: /node_modules/
        },
        {
          test: /\.(sass|less|css)$/,
          use: ['style-loader', 'css-loader', 'less-loader']
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        }
      ]
    },
    plugins: [
      new webpack.ProvidePlugin({
        Buffer: ['buffer', 'Buffer'],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        filename: "onsend.html",
        template: "./src/onsend/onsend.html",
        chunks: ["onsend"]
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/onsend/dialog.html",
        chunks: ["dialog"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: './assets',
            to: 'assets',
          },
          {
            from: dev ? 'manifest.dev.xml' : 'manifest.prod.xml',
            to: 'manifest.xml',
          }
        ]
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"]
      })
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: true,
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
