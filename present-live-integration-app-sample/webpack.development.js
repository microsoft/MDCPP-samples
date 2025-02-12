const Path = require('path');

const PORT = 8080;

module.exports = () => ({
  devtool: 'inline-source-map',
  devServer: {
    static: [Path.resolve(__dirname, 'dist'), Path.resolve(__dirname, 'src/static')],
    allowedHosts: ['localhost'],
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
      "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization"
    },
    port: PORT,
  },
});
