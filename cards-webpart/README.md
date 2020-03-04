![# Digital Assistant Cards for SharePoint](https://www.adenin.com/assets/images/identity/Logo_adenin.svg)

# Digital Assistant Cards for SharePoint

This client-side web part for SharePoint Online embeds Digital Assistant Cards into your SharePoint modern experience pages.

Cards can be displayed as persistent stand-alone Cards, or can be used in conjunction with the [PnP Modern Search Solution](https://github.com/microsoft-search/pnp-modern-search) to show relevant Cards in search results.

### Used SharePoint Framework Version 

![drop](https://img.shields.io/badge/version-1.10.0-blue.svg)

## Versioning

Versioning is managed manually. When deploying a new version, ensure the version is updated using semantic versioning.

The file containing version information is located at `cards-webpart/config/package-solution.json`.

## Preview the webpart

* Clone this repository
* In the `cards-webpart` project, run the following commands:
* `npm install`
* `gulp serve`
  
## Build the webpart for deployment

* Clone this repository
* In the `cards-webpart` project, run the following commands:
* `npm install`
* `gulp bundle --ship`
* `gulp package-solution --ship`

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN
* sharepoint/solution/cards-webpart.sppkg - the SPFx package to upload to AppCatalog