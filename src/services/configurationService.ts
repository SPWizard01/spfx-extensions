import { addOrUpdateExtensionConfigs } from "@spfx-extensions/core/idb";
import { CONFIGURATION_LIST_NAME, SPFXPREFIX } from "../utilities/constants";
import { ConfigurationListData, getCoreDefaultConfiguration } from "@spfx-extensions/core";


let appCatalogPromiseResolver = (_data: string | PromiseLike<string>) => {};
let appCatalogUrlPromise: Promise<string> | undefined; 

export async function getAppCatalogUrl(baseUrl = "") {
    if (appCatalogUrlPromise) {
        return appCatalogUrlPromise;
    }
    appCatalogUrlPromise = new Promise<string>((resolve) => {
        appCatalogPromiseResolver = resolve;
    });
    try {
        const apiResponse = await fetch(`${baseUrl}/_api/SP_TenantSettings_Current`, {
            headers: {
                Accept: "application/json;odata=verbose",
            }
        })
        const responseData = await apiResponse.json()
        const url = responseData.d.CorporateCatalogUrl as string;
        appCatalogPromiseResolver(url);
    }
    catch (err) {
        console.error(SPFXPREFIX, "Error while getting app catalog url. Trying default /sites/appcatalog", err);
        const fallBackUrl = `${window.location.origin}/sites/appcatalog`;
        appCatalogPromiseResolver(fallBackUrl);
    }
    return appCatalogUrlPromise;
}

async function ensureConfigurationListDataField(digestValue: string) {
    // /sites/appcatalog/_api/web/lists/GetByTitle('SPFxExtensionsConfiguration')/fields
    const appCatalogUrl = await getAppCatalogUrl();
    const fieldsUrl = `${appCatalogUrl}/_api/web/lists/GetByTitle('${CONFIGURATION_LIST_NAME}')/fields`;
    try {
        const req = await fetch(
            fieldsUrl,
            {
                headers: {
                    Accept: "application/json;odata=verbose",
                },
            }
        );
        if (req.status === 200) {
            const data = await req.json();
            const fields = data.d.results;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const fieldNames = fields.map((f: any) => f.InternalName);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const titleField = fields.find((f: any) => f.InternalName === "Title");
            if (!titleField) {
                console.error(SPFXPREFIX, "Title field not found.");
                return;
            }
            if (!titleField.EnforceUniqueValues) {
                // Update the Title field
                const updateFieldReq = await fetch(
                    `${appCatalogUrl}/_api/web/lists/GetByTitle('${CONFIGURATION_LIST_NAME}')/fields('${titleField.Id}')`,
                    {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digestValue,
                            "X-HTTP-Method": "MERGE",
                            "If-Match": "*",
                        },
                        body: JSON.stringify({
                            __metadata: {
                                type: "SP.Field",
                            },
                            Indexed: true,
                            EnforceUniqueValues: true,
                        }),
                    }
                );
                if (updateFieldReq.status === 204) {
                    console.info(SPFXPREFIX, "Title field updated successfully.");
                } else {
                    console.error(SPFXPREFIX, "Unable to update Title field.");
                }
            }
            if (!fieldNames.includes("Data")) {
                // Add the Data field
                const addFieldReq = await fetch(
                    fieldsUrl,
                    {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digestValue,
                        },
                        body: JSON.stringify({
                            __metadata: {
                                type: "SP.Field",
                            },
                            Title: "Data",
                            FieldTypeKind: 2,
                            Required: true,
                        }),
                    }
                );
                if (addFieldReq.status === 201) {
                    console.info(SPFXPREFIX, "Data field added successfully.");
                } else {
                    console.error(SPFXPREFIX, "Unable to add Data field.");
                }
            }
        }
    }
    catch (err) {
        console.error(SPFXPREFIX, "Error while ensuring configuration list data fields.", err);
    }
}

export async function ensureConfigurationList() {
    // /sites/appcatalog/_api/web/lists/GetByTitle('SPFxExtensionsConfiguration')
    const appCatalogUrl = await getAppCatalogUrl();
    try {
        const req = await fetch(
            `${appCatalogUrl}/_api/web/lists/GetByTitle('${CONFIGURATION_LIST_NAME}')`
        );
        const dgst = await getDigest();
        let newList = false;
        if (req.status === 404) {
            newList = true;
            console.log(SPFXPREFIX, "Creating configuration list.", dgst);
            // Create the list
            const createReq = await fetch(
                "/sites/appcatalog/_api/web/lists",
                {
                    method: "POST",
                    headers: {
                        Accept: "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": dgst,
                    },
                    body: JSON.stringify({
                        "__metadata": {
                            type: "SP.List",
                        },
                        // AllowContentTypes: false,
                        // ContentTypesEnabled: false,
                        BaseTemplate: 100,
                        Title: "SPFxExtensionsConfiguration",
                        Description: "Configuration list for SPFxExtensions",
                    }),
                }
            );
            if (createReq.status === 201) {
                console.info("Configuration list created successfully.");

            } else {
                console.error("Unable to create configuration list.");
            }
        }
        await ensureConfigurationListDataField(dgst);
        if (newList) {
            await ensureDefaultConfigurationListData();
        }
    }
    catch (err) {
        console.error(SPFXPREFIX, "Error while ensuring configuration list.", err);
    }
}



export async function getConfigurationListData() {
    const appCatalogUrl = await getAppCatalogUrl();
    const requestUrl = `${appCatalogUrl}/_api/web/lists/GetByTitle('${CONFIGURATION_LIST_NAME}')/items?$select=Title,Data`;
    const config = {
        headers: {
            Accept: "application/json;odata=verbose",
        }
    }
    let req = await fetch(
        requestUrl, {
        ...config
    }
    );
    if (req.status !== 200) {
        console.error(SPFXPREFIX, "Unable to fetch configuration list items.");
        return getCoreDefaultConfiguration(appCatalogUrl).map((c) => ({ ...c, date: new Date().toISOString(), expires: new Date().toISOString() }));
    }
    let data = await req.json();
    let results = data.d.results as ConfigurationListData[];
    let defaultAdded = false;
    if (results.length === 0) {
        await ensureDefaultConfigurationListData();
        defaultAdded = true;
    }
    if (defaultAdded) {
        req = await fetch(
            requestUrl, {
            ...config
        });
        data = await req.json();
        results = data.d.results as ConfigurationListData[];
    }
    await addOrUpdateExtensionConfigs(results);
    return results;
}

async function ensureDefaultConfigurationListData() {
    const dgst = await getDigest();
    const appCatalogUrl = await getAppCatalogUrl();

    for (const item of getCoreDefaultConfiguration(appCatalogUrl)) {
        const addReq = await fetch(
            `${appCatalogUrl}/_api/web/lists/GetByTitle('${CONFIGURATION_LIST_NAME}')/items`, {
            method: "POST",
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": dgst,
            },
            body: JSON.stringify({
                "__metadata": {
                    type: "SP.Data.SPFxExtensionsConfigurationListItem"
                },
                ...item,
            })
        }
        );
        if (addReq.status === 201) {
            console.info(SPFXPREFIX, `Item ${item.Title} added successfully.`);
        } else {
            console.error(SPFXPREFIX, `Unable to add item ${item.Title}.`);
        }
    }
}

async function getDigest() {
    const appCatalogUrl = await getAppCatalogUrl();
    const req = await fetch(
        `${appCatalogUrl}/_api/contextinfo`,
        {
            method: "POST",
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json",
            },
        }
    );
    if (req.status === 200) {
        const data = await req.json();
        return data.d.GetContextWebInformation.FormDigestValue;
    }
    return "";
}
