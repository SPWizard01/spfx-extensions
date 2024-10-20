import { DBSchema, openDB } from "idb";
import { ConfigurationListData } from "spfx-extensions-core";

import { SPFXPREFIX } from "../utilities/constants";

const Config = "Config";
interface SPFxExtensionsDB extends DBSchema {
    Config: {
        key: string;
        value: ConfigurationListData;
    };
}

const DBNAME = `SPFXEXT`;
const openDBPromise = openDB<SPFxExtensionsDB>(DBNAME, 1, {
    blocking(_currentVersion, _blockedVersion, _event) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        spfxDB.close();
        alert("A new version of this page is ready. Please reload the page.");
    },
    async upgrade(database, oldVersion, newVersion, transaction, event) {
        /// Create the object store
        if (oldVersion === 0) {
            database.createObjectStore(Config, { keyPath: "Title" });
        }
        //diff between 0 and 1 just delete the old database and let it be repopulated
        // if (oldVersion === 1) {
        //   database.createObjectStore(AllowedApps, { keyPath: "Id" });
        // }
    },
});
const spfxDB = await openDBPromise;
function getCacheItemBase(cacheTimeMinutes: number) {
    const dateNow = new Date();
    const date = dateNow.toISOString();
    dateNow.setMinutes(dateNow.getMinutes() + cacheTimeMinutes);
    const expires = dateNow.toISOString();
    return {
        date,
        expires,
    };
}
export async function getAllConfiguration() {
    await evictCache();
    return spfxDB.getAll(Config);
}
export async function getConfigByName(title: string) {
    await evictCache();
    return spfxDB.get(Config, title);
}
export async function addOrUpdateDataToCache(
    item: ConfigurationListData,
    cacheTimeMinutes = 60
) {
    await spfxDB.put(Config, {
        ...item,
        ...getCacheItemBase(cacheTimeMinutes),
    });
}
export function addOrUpdateManyToCache(items: ConfigurationListData[]) {
    const tx = spfxDB.transaction(Config, "readwrite");
    const txStore = tx.objectStore(Config);
    items.forEach((u) => txStore.put({ ...u, ...getCacheItemBase(60) }));
    return tx.done;
}

export function removeConfigFromCache(title: string) {
    return spfxDB.delete(Config, title);
}
export function removeManyFromCache(title: string[]) {
    const tx = spfxDB.transaction(Config, "readwrite");
    const txStore = tx.objectStore(Config);
    title.forEach((u) => txStore.delete(u));
    return tx.done;
}
export async function evictCache() {
    const cache = await spfxDB.getAll(Config);
    const nowTime = new Date();
    const cacheToRemove = cache.filter((ci) => {
        const itemExpires = new Date(ci.expires);
        //the items that should be removed
        return nowTime >= itemExpires;
    });

    const toEvict = cacheToRemove.length;
    if (toEvict > 0) {
        await removeManyFromCache(cacheToRemove.map((c) => c.Title));
        console.warn(SPFXPREFIX, `Evicted ${toEvict} items from ${Config} cache.`);
    }
}
