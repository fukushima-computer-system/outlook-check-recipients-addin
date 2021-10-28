param($Mode)

Copy-Item $PSScriptRoot/manifest.$($Mode).xml $PSScriptRoot/manifest.xml
Copy-Item $PSScriptRoot/env.$($Mode).json $PSScriptRoot/env.json

if ( $Mode -eq "dev" ) {
    office-addin-debugging start manifest.xml web
}
if ( $Mode -eq "prod" ) {
    webpack --mode production
}
