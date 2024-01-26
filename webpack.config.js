var project = require("./package.json");
var path = require("path");
const MomentLocalesPlugin = require('moment-locales-webpack-plugin');
const MomentTimezoneDataPlugin = require('moment-timezone-data-webpack-plugin');
const currentYear = new Date().getFullYear();

// Return the configuration
module.exports = (env, argv) => {
    var isDev = argv.mode !== "production";
    return {
        // Set the main source as the entry point
        entry: [
            path.resolve(__dirname, project.main)
        ],

        // Output location
        output: {
            path: path.resolve(__dirname, "dist"),
            filename: project.name + (isDev ? "" : ".min") + ".js",
            publicPath: ""
        },

        plugins: [
            // Only keep 'en' locales with Moment by default
            new MomentLocalesPlugin(),

            // Only keep Moment timezone data for the past year and the next 3 years matching common areas in America
            new MomentTimezoneDataPlugin({
                matchZones: [
                    'America/Anchorage', 'America/Chicago', 'America/Denver', 'America/Los_Angeles',
                    'America/New_York', 'America/Phoenix', 'America/Puerto_Rico', 'Pacific/Guam',
                    'Pacific/Honolulu'
                ],
                startYear: currentYear - 1,
                endYear: currentYear + 3
            }),
        ],
        
        // Resolve the file names
        resolve: {
            extensions: [".js", ".css", ".scss", ".ts"]
        },

        // Dev Server
        devServer: {
            inline: true,
            hot: true,
            open: true,
            publicPath: "/dist/"
        },

        // Loaders
        module: {
            rules: [
                // SASS to JavaScript
                {
                    // Target the sass and css files
                    test: /\.s?css$/,
                    // Define the compiler to use
                    use: [
                        // Create style nodes from the CommonJS code
                        { loader: "style-loader" },
                        // Translate css to CommonJS
                        { loader: "css-loader" },
                        // Compile sass to css
                        {
                            loader: "sass-loader",
                            options: {
                                implementation: require("sass")
                            }
                        }
                    ]
                },
                // Handle Image Files
                {
                    test: /\.(jpe?g|png|gif|svg|eot|woff|ttf)$/,
                    loader: "url-loader"
                },
                // JavaScript
                {
                    // Target JavaScript files
                    test: /\.jsx?$/,
                    use: [
                        // JavaScript (ES5) -> JavaScript (Current)
                        {
                            loader: "babel-loader",
                            options: { presets: ["@babel/preset-env"] }
                        }
                    ]
                },
                // TypeScript to JavaScript
                {
                    // Target TypeScript files
                    test: /\.tsx?$/,
                    use: [
                        // JavaScript (ES5) -> JavaScript (Current)
                        {
                            loader: "babel-loader",
                            options: { presets: ["@babel/preset-env"] }
                        },
                        // TypeScript -> JavaScript (ES5)
                        { loader: "ts-loader" }
                    ]
                }
            ]
        }
    };
}