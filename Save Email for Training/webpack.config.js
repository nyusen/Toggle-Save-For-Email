/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const env = process.env;  // Access environment variables

const dev = env.NODE_ENV === "development";

const urlDev = "https://localhost:3000/";
const urlProd = "https://localhost:3000/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: "./src/commands.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      clean: true,
      library: {
        name: 'commands',
        type: 'window',
      },
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            options: {
              compilerOptions: {
                noEmit: false,
              },
            },
          },
        },
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog.html",
        chunks: ["polyfill", "dialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "signin-dialog.html",
        template: "./src/signin-dialog.html",
        chunks: ["polyfill"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth-callback.html",
        template: "./src/auth-callback.html",
        chunks: ["polyfill"],
      }),
      new HtmlWebpackPlugin({
        filename: "sign-in-start.html",
        template: "./src/sign-in-start.html",
        chunks: ["polyfill"],
      }),
      new HtmlWebpackPlugin({
        filename: "retry-dialog.html",
        template: "./src/retry-dialog.html",
        chunks: ["polyfill"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/",
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: dev ? "https" : "http",
        options: dev ? await getHttpsOptions() : {},
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      allowedHosts: 'all',
    },
  };

  return config;
};
