const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const mode = argv.mode || "development";
  const isProd = mode === "production";

  // Get trusted dev certs for HTTPS (installs to trusted root on first run)
  const httpsOptions = isProd ? {} : await devCerts.getHttpsServerOptions();

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [
            isProd ? MiniCssExtractPlugin.loader : "style-loader",
            "css-loader",
          ],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "privacy.html", to: "privacy.html" },
          { from: "support.html", to: "support.html" },
          { from: "index.html", to: "index.html" },
        ],
      }),
      ...(isProd
        ? [new MiniCssExtractPlugin({ filename: "[name].css" })]
        : []),
    ],
    devServer: {
      static: {
        directory: path.resolve(__dirname, "dist"),
      },
      port: 3000,
      hot: false,
      server: {
        type: "https",
        options: httpsOptions,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      allowedHosts: "all",
    },
    devtool: isProd ? false : "source-map",
  };
};
