const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const isProduction = argv.mode === "production";

  const config = {
    mode: isProduction ? "production" : "development",
    devtool: isProduction ? false : "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets" }],
      }),
    ],
  };

  // Dev-only: HTTPS via office-addin-dev-certs + a proxy to the Python
  // backend. A production build has no dev server at all — server.py
  // serves these built files directly (see Part B of the packaging plan).
  if (!isProduction) {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    config.devServer = {
      static: path.resolve(__dirname, "dist"),
      server: {
        type: "https",
        options: httpsOptions,
      },
      port: 3000,
      proxy: [
        {
          // server.py owns the /api/* prefix directly — forward unchanged,
          // don't strip it.
          context: ["/api"],
          target: "http://127.0.0.1:8765",
          changeOrigin: true,
        },
      ],
    };
  }

  return config;
};
