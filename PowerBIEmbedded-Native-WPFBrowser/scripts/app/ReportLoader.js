/// <reference path="../powerbi.js" />

function LoadEmbeddedObject(embedUrl, accessToken, embedId, embedType, tokenType, dashboardId) {

    var models = window['powerbi-client'].models
    var permissions = models.Permissions.All;
    
    var config = {
        type: embedType,
        accessToken: accessToken,
        tokenType: tokenType,
        embedUrl: embedUrl,
        id: embedId,
        dashboardId: dashboardId,
        permissions: permissions,
        settings: {
            filterPaneEnabled: true,
            navContentPaneEnabled: true
        }
    };


    // Grab the reference to the div HTML element that will host the report.
    var embedContainer = document.getElementById('EmbedContainer');
    // Embed the report and display it within the div container.
    var embed = powerbi.embed(embedContainer, config);

    // Report.off removes a given event handler if it exists.
    embed.off("loaded");

    // Report.on will add an event handler which prints to Log window.

    embed.on("loaded", function () {
        window.external.LogToBrowserHost("Loaded");
        // Report.off removes a given event handler if it exists.

        embed.off("loaded");

    });

    embed.on("error", function (event) {
        window.external.LogToBrowserHost(event.detail);

        embed.off("error");
    });

    // Report.on will add an event handler which prints to Log window.

    embed.on("rendered", function () {
        window.external.LogToBrowserHost("Rendered");
        // Report.off removes a given event handler if it exists.
        embed.off("rendered");

    });

    embed.off("saved");
    embed.on("saved", function (event) {
        window.external.LogToBrowserHost(event.detail);
        if (event.detail.saveAs) {
            window.external.LogToBrowserHost('In order to interact with the new report, create a new token and load the new report');
        }
    });
    if (embedType === "tile") {
        embed.off("tileLoaded");
        //// Tile.on will add an event handler which prints to Log window.
        embed.on("tileLoaded", function (event) {
            window.external.LogToBrowserHost("Tile loaded event");
            window.external.LogToBrowserHost(JSON.stringify(event.detail));
        });
    };
    if (embedType === "dashboard" || embedType === "tile") {
        embed.off("tileClicked");
        embed.on("tileClicked", function (event) {
            window.external.LogToBrowserHost(JSON.stringify(event.detail));
        });
    };

};