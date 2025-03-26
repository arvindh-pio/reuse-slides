import { APIS } from "../utils/api";
import { getToken } from "../utils/utils";
import useFiles from "./useFiles";

interface IFilter {
    name: string;
    type: string;
}

interface IConfig {
    siteName: string;
    libraryName: string;
    filters: IFilter[];
}

const useInitial = () => {
    const token = getToken();
    const { getLibraryId, customFetchFilesWithThumbnails } = useFiles();

    // 1 -> given a site name and library name, we need to fetch
    const fetchPPTFiles = async (config: IConfig) => {
        try {
            const res = await fetchCustomFilesFromOtherSites(config);
            const filterConfigs = await fetchFilters(res?.site?.id, config);

            return {
                files: res?.data,
                site: res?.site,
                drive: res?.drive,
                filterConfigs
            };
        } catch (error) {
            console.log("Error fetching ppt files ", error);
            return {};
        }
    }

    const fetchFilters = async (localSiteId: string, config: IConfig) => {
        const { libraryName, filters } = config;
        const token = getToken();
        const library = await getLibraryId(localSiteId, libraryName);

        try {
            const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${localSiteId}/lists/${library.id}/columns`, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                method: "GET"
            });
            const colsData = await response.json();
            let cols = colsData?.value?.filter((col) => {
                const data = filters?.find((filter) => filter?.name === col?.displayName);
                return data !== undefined ? col : null;
            });
            cols = cols?.filter((col) => col !== null);
            return cols;
        } catch (error) {
            console.log("filters error -> ", error);
        }
    }

    const fetchCustomFilesFromOtherSites = async (config: IConfig) => {
        const sitesResponse = await fetch(
            `https://graph.microsoft.com/v1.0/sites/pixerenet1.sharepoint.com:/sites/${config?.siteName}`, {
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            method: "GET"
        })
        const site = await sitesResponse?.json();

        const drivesResponse = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${site?.id}/drives`, {
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            method: "GET"
        })
        const drives = await drivesResponse?.json();
        const findDrive = drives?.value?.find((drive) => drive?.name === config?.libraryName);

        const finalResponse = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${site?.id}/drives/${findDrive?.id}/root/search(q='.pptx')`, {
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            method: "GET"
        })
        const finalData = await finalResponse?.json();

        const dd = await customFetchFilesWithThumbnails(site?.id, config?.libraryName, finalData?.value);
        return {
            site: site,
            drive: findDrive,
            data: dd
        };
    }

    return { fetchPPTFiles };
}

export default useInitial;