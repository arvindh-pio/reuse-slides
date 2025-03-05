import { APIS } from "../utils/api";
import { getToken } from "../utils/utils";
import useFiles from "./useFiles";

interface IConfig {
    siteName: string;
    libraryName: string;
}

const useInitial = () => {
    const token = getToken();
    const { fetchFiles } = useFiles();

    const fetchPPTFiles = async (config: IConfig) => {
        try {
            const site = await fetchUserSite(config?.siteName);
            const documentLibrary = await fetchDocumentLibrary(site?.id, config?.libraryName);
            const files = await fetchFiles(documentLibrary?.id, site?.id, config?.libraryName);
            return {
                files,
                site, 
                drive: documentLibrary
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

    return { fetchPPTFiles };
}

export default useInitial;