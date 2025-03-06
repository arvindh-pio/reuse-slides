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
    const { fetchFiles, getLibraryId } = useFiles();

    const fetchPPTFiles = async (config: IConfig) => {
        try {
            const site = await fetchUserSite(config?.siteName);
            const documentLibrary = await fetchDocumentLibrary(site?.id, config?.libraryName);
            const files = await fetchFiles(documentLibrary?.id, site?.id, config?.libraryName);
            const filterConfigs = await fetchFilters(site?.id, config);

            return {
                files,
                site,
                drive: documentLibrary,
                filterConfigs
            };
        } catch (error) {
            console.log("Error fetching ppt files ", error);
            return {};
        }
    }

    const fetchUserSite = async (siteName: string) => {
        try {
            // first get all sites
            const sitesResponse = await fetch(APIS.GET_ALL_SITES, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                method: "GET"
            })
            const sitesData = await sitesResponse.json();
            const site = sitesData?.value?.find((site) => site?.displayName === siteName);
            return site;
            // setSite(site);🔴
            // await getDocAndPPTFiles(site);
        } catch (error) {
            console.log("Error in searching sites ", error);
        } finally {
            // setLoading(false);🔴
        }
    }

    const fetchDocumentLibrary = async (siteId: string, libraryName: string) => {
        try {
            const drivesResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
                method: "GET",
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
            });
            const drivesData = await drivesResponse.json();
            const drive = drivesData?.value?.find(drive => drive.name.toLowerCase() === libraryName.toLowerCase());
            return drive;
        } catch (error) {
            console.log("Error while searching for doc in drives", error);
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

    return { fetchPPTFiles };
}

export default useInitial;