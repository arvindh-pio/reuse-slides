import React, { useEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import FileInput from "./FileInput";
import Ppt from "../Pages/Ppt";
import Slides from "../Pages/Slides";
import { CustomDriveItemResponse } from "../Types";
import { API_BASE_URL, UPLOAD_API } from "../utils/constants";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { getToken } from "../utils/utils";
import Config from "./Config";
import FilterDropdown from "./FilterDropdown";
import { getFilteredData } from "../utils/filters";

export interface ISlide {
  index: number;
  slideId: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "1rem"
  },
  searchDiv: {
    display: "flex",
    gap: "5px",
    alignItems: "center",
    width: "100%",
    boxSizing: "border-box",
    marginBottom: "0.7rem",
    flexDirection: "column"
  },
  searchInput: {
    padding: "5px 10px",
    fontSize: "0.9rem",
    borderRadius: "5px",
    outline: "none",
    border: "1px solid black",
    boxSizing: "border-box",
    width: "100%",
  },
  actionBtns: {
    display: "flex",
    width: "100%",
    justifyContent: "space-evenly"
  },
  searchButton: {
    background: "none",
    border: "1px solid crimson",
    padding: "5px 10px",
    fontFamily: "sans-serif",
    borderRadius: "4px",
    cursor: "pointer"
  },
  browseButton: {
    background: "none",
    border: "1px solid crimson",
    padding: "10px 20px",
    cursor: "pointer",
    fontFamily: "sans-serif",
    fontSize: "0.9rem",
    textAlign: "center",
    borderRadius: "8px",
    transition: "all 0.3s",
    backgroundColor: "crimson",
    color: "white",
    fontWeight: "bold",
    letterSpacing: "0.1em",
    "&:hover": {
      opacity: "0.7",
    }
  },
  block: {
    display: "block",
    margin: "0"
  },
  loading: {
    width: "100%",
    height: "75vh",
    display: "flex",
    justifyContent: "center",
    alignItems: "center"
  }
});

const App: React.FC = () => {
  // styles
  const styles = useStyles();
  const [searchQuery, setSearchQuery] = useState("");

  // states
  const [file, setFile] = useState<File | null>(null);

  const [previews, setPreviews] = useState<string[]>([]);
  const [sourceSlideIds, setSourceSlideIds] = useState<ISlide[]>([]);
  const [base64, setBase64] = useState<string | null>(null);
  const [formatting, setFormatting] = useState(true);

  const [recentResults, setRecentResults] = useState<CustomDriveItemResponse[]>([]);
  const [searchResults, setSearchResults] = useState<CustomDriveItemResponse[]>([]);

  const [showSlides, setShowSlides] = useState(false);
  const [isSearchClicked, setSearchClicked] = useState(false);

  const [libraryName, setLibraryName] = useState("");
  const [siteName, setSiteName] = useState("");

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [site, setSite] = useState(null);
  const [drive, setDrive] = useState(null);
  const [config, setConfig] = useState(false);

  // filter columns
  const [tag, setTag] = useState(null);

  const callGraphApi = async (type: "RECENT" | "SEARCH") => {
    if (type === "SEARCH") {
      setError(null);
      setSearchClicked(true);
    }

    const token = localStorage.getItem("token");
    const url = `https://graph.microsoft.com/v1.0/search/query`;
    const reqBody = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: {
            queryString: type === "RECENT" ? "filetype:pptx OR filetype.ppt" : `${searchQuery} AND filetype:pptx`
          }
        }
      ]
    }

    const response = await
      fetch(url,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify(reqBody)
        }
      );

    const data = await response.json();
    const pptFiles = await fetchThumbnails(data?.value?.[0]?.hitsContainers?.[0]?.hits);

    if (type === "RECENT") {
      setRecentResults(pptFiles.slice(0, 7));
    } else {
      setSearchResults(pptFiles);
    }
  }

  const searchPpt = async () => {
    await callGraphApi("SEARCH");
  }

  const generatePreviews = async () => {
    try {
      setShowSlides(true);
      setLoading(true);

      const formData = new FormData();
      formData.append("ppt", file);

      const response = await fetch(`${API_BASE_URL}${UPLOAD_API}`, {
        method: 'POST',
        body: formData,
        headers: {
          "Content-Type": "multipart/form-data"
        }
      });

      if (!response.ok) {
        throw new Error('Failed to upload file');
      }

      const data = await response.json();
      setPreviews(data.slides);
    } catch (error) {
      console.log("Generate previews Error -> ", error);
      setError(error.message);
    } finally {
      setLoading(false);
    }
  };

  const extractSlideIds = async (file: File): Promise<ISlide[]> => {
    try {
      const zip = new JSZip();
      const pptx = await zip.loadAsync(file);

      const xmlData = await pptx.file("ppt/presentation.xml").async("text");

      const parser = new XMLParser({ ignoreAttributes: false });
      const parsedXml = parser.parse(xmlData);

      const slidesList = parsedXml?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"];

      if (!slidesList) {
        return [];
      }

      const slidesArray = Array.isArray(slidesList) ? slidesList : [slidesList];
      const slideDataPromises = slidesArray?.map(async (slide, index) => {
        return {
          index: index + 1,
          slideId: slide["@_id"] || `Unknown_${index + 1}`
        }
      })

      const slideData = await Promise.all(slideDataPromises);
      setSourceSlideIds(slideData);
      return slideData;
    } catch (error) {
      console.log("Error parsing PPTX: ", error);
      return [];
    }
  };

  const getBase64 = async (file: File) => {
    return new Promise((resolve, reject) => {
      try {
        const reader = new FileReader();

        reader.onload = async (_) => {
          const startIndex = reader.result.toString().indexOf("base64,");
          const copyBase64 = reader.result.toString().slice(startIndex + 7);

          resolve(copyBase64);
        };

        reader.readAsDataURL(file);
      } catch (error) {
        console.log("err -> ", error);
        reject(error?.message);
      }
    })
  };

  const getFile = async (file: CustomDriveItemResponse) => {
    const token = localStorage.getItem("token");
    const response = await
      fetch(`https://graph.microsoft.com/v1.0/drives/${file?.parentReference?.driveId}/items/${file?.id}/content`,
        {
          method: "GET",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
        }
      );
    const data = await response.blob();
    return new File([data], file?.name, { type: data.type });
  }

  const generatePPTDetails = async (file: File | CustomDriveItemResponse) => {
    if (file instanceof File) {
      const slideIds = await extractSlideIds(file);
      setSourceSlideIds(slideIds);
      const base64 = await getBase64(file);
      setBase64(base64 as string);
      generatePreviews();
    } else {
      const onlineFile = await getFile(file);
      const slideIds = await extractSlideIds(onlineFile);
      setSourceSlideIds(slideIds);
      const base64 = await getBase64(onlineFile);
      setBase64(base64 as string);
      generatePreviews();
    }
  }

  const fetchThumbnails = async (data) => {
    const token = localStorage.getItem("token");

    const pptFiles = await Promise.all(
      data.map(async (hit) => {
        const file = hit;
        const extArr = file.name.split(".") || [];
        const ext = extArr[extArr.length - 1];

        if (ext === "ppt" || ext === "pptx") {
          let image = null;
          try {
            const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/thumbnails`, {
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

          // # fallback 1
          if (!image) {
            const siteId = file?.parentReference?.siteId;
            try {
              const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${file.id}/thumbnails`, {
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
          }

          // # fallback 2
          if (!image) {
            const siteId = file?.parentReference?.siteId;
            try {
              const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/preview`, {
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
          }

          return {
            thumbnail: image,
            ...file
          };
        }
        return null; // Return null for non-PPT files
      })
    );

    return pptFiles;
  }

  // fetches whole onedrive and sharepoint recent ppt files
  const fetchFiles = async () => {
    await callGraphApi("RECENT");
  }

  const handleBack = () => {
    // switch to initial states
    setSearchResults([]);
    setSearchQuery("");
    setShowSlides(false);
    setSearchClicked(false);
    setError("");

    setPreviews([]);
    setBase64(null);
    setSourceSlideIds([]);
  }

  // 1. after setting config, fetch the recent files only from that config
  const getRecentFilesFromLibrary = async () => {
    setLoading(true);
    setConfig(true);
    const token = getToken();

    try {
      // first get all sites
      const sitesResponse = await fetch("https://graph.microsoft.com/v1.0/sites?search=*", {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        method: "GET"
      })
      const sitesData = await sitesResponse.json();
      const site = sitesData?.value?.find((site) => site?.displayName === siteName);
      setSite(site);
      await getDocAndPPTFiles(site);
    } catch (error) {
      console.log("Error in searching files ", error);
    } finally {
      setLoading(false);
    }
  }

  // 2. gets library and ppt files from that lib
  const getDocAndPPTFiles = async (localSite) => {
    const token = getToken();
    const siteId = localSite.id;

    const drive = await getDocLibraryFromSite(siteId, token);

    if (drive) {
      setDrive(drive);
      await getFilesFromDrive(siteId, drive.id, token);
    } else {
      console.log("Drive not found ", libraryName);
    }
  }

  // 2.1 get particular library from site
  const getDocLibraryFromSite = async (siteId: string, token: string) => {
    try {
      const drivesResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
      })
      const drivesData = await drivesResponse.json();

      const drive = drivesData?.value?.find(drive => drive.name.toLowerCase() === libraryName.toLowerCase());
      return drive;
    } catch (error) {
      console.log("Error while searching for doc in drives", error);
    }
  }

  // 2.2 get ppt files from drive
  const getFilesFromDrive = async (siteId: string, driveId: string, token: string) => {
    try {
      const defaultFiles = await getFiles(driveId, token);
      const expandedFiles = await getExpandedFiles(siteId);
      const files = mergeResponses(defaultFiles, expandedFiles);
      const ppts = await getThumbnails(files, driveId);
      setRecentResults(ppts);
    } catch (error) {
      console.log("error in getting files ", error);
    }
    // try {
    //   const library = await getLibraryId(siteId);
    //   const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
    //   // const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${library.id}/items?$expand=fields`;
    //   const driveResponse = await fetch(url, {
    //     headers: {
    //       Authorization: `Bearer ${token}`,
    //       "Content-Type": "application/json"
    //     },
    //     method: "GET"
    //   });
    //   const drivesData = await driveResponse.json();
    //   const pptFiles = drivesData?.value?.filter(file => file?.name?.endsWith(".pptx"));
    //   // const pptsPromises = pptFiles?.map(async (ppt) => {
    //   //   const id = ppt.id;

    //   //   try {
    //   //     const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${library.id}/files/${id}?$expand=fields`;
    //   //     const response = await fetch(url, {
    //   //       headers: {
    //   //         Authorization: `Bearer ${token}`,
    //   //         "Content-Type": "application/json"
    //   //       },
    //   //       method: "GET"
    //   //     });
    //   //     const data = await response.json();
    //   //     return {
    //   //       customFileId: id,
    //   //       ...data.value
    //   //     }
    //   //   } catch (error) {

    //   //   }
    //   // })
    //   // const ff = await Promise.all(pptsPromises);
    //   const ppts = await getThumbnails(pptFiles, driveId);
    //   console.log("xxxxxxxx ", pptFiles);
    //   setRecentResults(ppts);
    // } catch (error) {
    //   console.log("error get files from drive -> ", error);
    // }
  }

  // 2.2.1 get default files
  const getFiles = async (driveId: string, token: string) => {
    try {
      const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
      const driveResponse = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        method: "GET"
      });
      const drivesData = await driveResponse.json();
      const pptFiles = drivesData?.value?.filter(file => file?.name?.endsWith(".pptx"));
      return pptFiles;
    } catch (error) {
      console.log("error get files from drive -> ", error);
    }
  }

  // 2.2.2 get expanded files
  const getExpandedFiles = async (siteId) => {
    const token = getToken();
    const library = await getLibraryId(siteId);

    try {
      const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId || site.id}/lists/${library.id}/items?$expand=fields`, {
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

  // 2.2.3 merge responses
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
    console.log("########", final);
    return final;
  }

  // get thumbnails
  const getThumbnails = async (ppts, _: string = null) => {
    const token = localStorage.getItem("token");

    const pptFiles = await Promise.all(
      ppts.map(async (hit) => {
        const file = hit;
        console.log("file ", file);
        const extArr = file?.fields?.FileLeafRef.split(".") || [];
        const ext = extArr[extArr.length - 1];

        if (ext === "ppt" || ext === "pptx") {
          let image = null;
          // # fallback 1
          const siteId = file?.parentReference?.siteId;
          const driveId = file?.customDriveId;
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
        return null; // Return null for non-PPT files
      })
    );

    return pptFiles;
  }

  const searchForKeywordInLibraryDocs = async () => {
    setError(null);
    setSearchClicked(true);
    const token = getToken();

    const url = `https://graph.microsoft.com/v1.0/sites/${site.id}/drives/${drive.id}/root/search(q='{${searchQuery}}')`;
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        }
      });
      const data = await response.json();
      const pptFiles = data?.value?.filter(file => file?.name?.endsWith(".pptx"));
      const ppts = await getThumbnails(pptFiles);
      setSearchResults(ppts);
    } catch (error) {
    }
  }

  const handleReset = () => {
    setSearchQuery("");
    setSearchResults([]);
    setSearchClicked(false);
    setError("");
  }

  const getLibraryId = async (localSiteId: string = null) => {
    const token = getToken();

    try {
      const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${localSiteId || site.id}/lists`, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        method: "GET"
      });
      const libraryData = await response.json();
      console.log(libraryData, "********");
      const library = libraryData?.value?.find((data) => data.displayName === libraryName);
      return library;
    } catch (error) {
      console.log("filters error -> ", error);
    }
  }

  const getFilters = async () => {
    const token = getToken();
    const library = await getLibraryId();

    try {
      const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${library.id}/columns`, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        method: "GET"
      });
      const colsData = await response.json();
      const tagCol = colsData?.value?.find((col) => col?.displayName === "Tag");
      setTag(tagCol);
    } catch (error) {
      console.log("filters error -> ", error);
    }
  }

  const handleFilter = async (key: string, val: string) => {
    const datas = getFilteredData(recentResults, key, val);
    setSearchResults(datas);
  }

  const test = async () => {
    const token = getToken();
    const library = await getLibraryId(site.id);

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${site.id}/drives/b!A8xRw8PhhEipKKGLOg8jIdLYbQIfU9pOn4DXZs7wMSE2CHN5BdKHRr8UO3rRbCpK/root/search(q='presentation')?$filter=fields/Tag eq 'Tech'`, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        method: "GET"
      });
      const colsData = await response.json();
      console.log(colsData, "^^^^^^");
    } catch (error) {
      console.log("filters error -> ", error);
    }
  }

  useEffect(() => {
    if (site && drive) {
      getFilters();
      test();
    }
  }, [site, drive])

  useEffect(() => {
    if (!showSlides) {
      setSearchQuery("");
    }
  }, [showSlides])

  if (loading) return <p className={styles.loading}>Loading...</p>

  return (
    <div className={styles.root}>
      {/* common */}
      {/* search, browse */}
      {!config ? (
        <Config
          libraryName={libraryName}
          setLibraryName={setLibraryName}
          siteName={siteName}
          setSiteName={setSiteName}
          getRecentFilesFromLibrary={getRecentFilesFromLibrary} />
      ) : (
        <>
          <div className={styles.searchDiv}>
            <input
              type="text"
              name="searchQuery"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className={styles.searchInput} />
            <div className={styles.actionBtns}>
              <button
                onClick={searchForKeywordInLibraryDocs}
                className={styles.searchButton}>Search</button>
              <button className={styles.searchButton} onClick={handleReset} disabled={!isSearchClicked}>Reset</button>
            </div>
          </div>
          {/* browse */}
          <FileInput
            setFile={setFile}
            generatePPTDetails={generatePPTDetails} />
          {error && <p className={styles.block}>{error}</p>}

          {/* Filters */}
          <h3>Filters: </h3>
          {tag && (
            <FilterDropdown customObject={tag} site={site} library={drive} handleFilter={handleFilter} />
          )}

          {!showSlides
            ? <Ppt
              searchResults={(searchResults?.length > 0 || isSearchClicked) ? searchResults : recentResults}
              generatePPTDetails={generatePPTDetails}
              isSearchClicked={isSearchClicked} />
            : loading ? (
              <p>Loading...</p>
            ) : (
              <Slides
                base64={base64}
                previews={previews}
                sourceSlideIds={sourceSlideIds}
                formatting={formatting}
                setFormatting={setFormatting}
                handleBack={handleBack} />
            )}
        </>
      )}
    </div>
  );
};

export default App;
