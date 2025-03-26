import { getToken } from "../utils/utils";

const useFiles = () => {
    const token = getToken();

    const customFetchFilesWithThumbnails = async (siteId: string, libraryName: string, existingFiles: any) => {
        try {
            const expandedFiles = await fetchExpandedFiles(siteId, libraryName);
            const files = mergeResponses(existingFiles, expandedFiles);
            const ppts = await getThumbnails(files);
            return ppts;
        } catch (error) {
            console.log("error in getting files ", error);
            return [];
        }
    }

    const fetchExpandedFiles = async (siteId: string, libraryName: string) => {
        const library = await getLibraryId(siteId, libraryName);

        try {
            const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${library.id}/items?$expand=fields`, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                method: "GET"
            });
            const libraryData = await response.json();
            return libraryData?.value;
        } catch (error) {
            console.log("get expanded files error -> ", error);
        }
    }

    const getLibraryId = async (localSiteId: string, libraryName: string) => {
        try {
            const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${localSiteId}/lists`, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                method: "GET"
            });
            const libraryData = await response.json();
            const library = libraryData?.value?.find((data) => data.displayName === libraryName);
            return library;
        } catch (error) {
            console.log("get library id error -> ", error);
        }
    }

    const mergeResponses = (defaultFiles, expandedFiles) => {
        let final = defaultFiles?.map((defFile) => {
            const found = expandedFiles?.find((file) => file?.fields?.FileLeafRef === defFile.name);
            if (found) return {
                customId: defFile?.id,
                customDriveId: defFile?.parentReference?.driveId,
                ...found
            };
            else null;
        });

        final = final?.filter((file) => file !== null);
        return final;
    }

    const getThumbnails = async (ppts) => {
        const pptFiles = await Promise.all(
            ppts.map(async (hit) => {
                const file = hit;
                const extArr = file?.fields?.FileLeafRef ? file?.fields?.FileLeafRef.split(".") : file?.name.split(".");
                const ext = extArr[extArr.length - 1];

                if (ext === "ppt" || ext === "pptx") {
                    let image = null;
                    // # fallback 1
                    const siteId = file?.parentReference?.siteId;
                    const driveId = file?.customDriveId || file?.parentReference?.driveId;
                    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${file.customId}/thumbnails`;

                    try {
                        const response = await fetch(url, {
                            method: "GET",
                            headers: {
                                Authorization: `Bearer ${token}`,
                                "Content-Type": "application/json"
                            }
                        });
                        const data = await response.json();
                        image = data?.value?.[0]?.large?.url;
                    } catch (error) {
                        image = null;
                    }

                    return {
                        thumbnail: image,
                        ...file
                    };
                }
                return null;
            })
        );

        return pptFiles;
    }

    return { getThumbnails, getLibraryId, customFetchFilesWithThumbnails };
}

export default useFiles;