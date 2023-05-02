## pnp - Extensibility Library Demo Project

This project shows how to implement custom data sources for the 'Search Results' Web Part.

![alt System News](https://github.com/PDSB-sps/SystemLeadershipExt/blob/main/src/screenshots/srch_ext.png)

### Documentation

A complete step-by-step tutorial is available [here](https://microsoft-search.github.io/pnp-modern-search/extensibility/)


### References for some issues:

- pnp-modern-search-extensibility-samples
https://github.com/microsoft-search/pnp-modern-search-extensibility-samples/tree/main/search-extensibility-samples

- WebpartContext in Custom web component not available - Search Extensibility Library
https://github.com/microsoft-search/pnp-modern-search/issues/918

- create context for SPHttpClient and PageContext using this._servicescope from BaseWebComponent while working with custom web component
https://github.com/microsoft-search/pnp-modern-search/issues/1738

- Does all attachments in list get crawled in SP Online?
https://sharepoint.stackexchange.com/questions/263678/does-all-attachments-in-list-get-crawled-in-sp-online

- Attachments Search O365
https://sharepoint.stackexchange.com/questions/208577/attachments-search-o365
https://www.techmikael.com/2014/04/solution-to-displaying-attachments-for.html

- Create Todo task
https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks?view=graph-rest-1.0&tabs=javascript

- Create plannerTask
https://learn.microsoft.com/en-us/graph/api/planner-post-tasks?view=graph-rest-1.0&tabs=http


### Issues on Adding the planner files - References:

- sp-dev-fx-webparts/samples/react-mytasks
https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-mytasks

- TSLint is not supported for rush-stack-compiler-4.X packages.
https://techcommunity.microsoft.com/t5/sharepoint-developer/tslint-is-not-supported-for-rush-stack-compiler-4-x-packages/m-p/3267711

- Calling MS Graph API: An error as soon as declaring the MSGraphClient
https://sharepoint.stackexchange.com/questions/257102/calling-ms-graph-api-an-error-as-soon-as-declaring-the-msgraphclient
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph

- Property 'web' does not exist on type 'SPRest'
https://sharepoint.stackexchange.com/questions/273921/property-web-does-not-exist-on-type-sprest

- PnP JS
https://pnp.github.io/pnpjs/getting-started/
https://pnp.github.io/pnpjs/packages/
https://pnp.github.io/pnpjs/graph/behaviors/

- Using the spfi
const sp = spfi(siteUrl).using(SPFx(context));

- PnP libs used
https://pnp.github.io/pnpjs/packages/#graph
https://pnp.github.io/pnpjs/graph/photos/
https://pnp.github.io/pnpjs/graph/planner/
https://pnp.github.io/pnpjs/graph/users/


### Planner task details update
- MS Graph API ... create Planner task including details/description using one single HTTP request?
https://stackoverflow.com/questions/75316467/ms-graph-api-create-planner-task-including-details-description-using-one-sin?rq=1
https://stackoverflow.com/questions/48851611/how-can-i-create-a-planner-task-with-a-description
https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0#common-planner-error-conditions



