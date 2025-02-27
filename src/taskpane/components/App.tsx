import React, { useEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import FileInput from "./FileInput";
import Ppt from "../Pages/Ppt";
import Slides from "../Pages/Slides";
import { CustomDriveItemResponse } from "../Types";
import { API_BASE_URL, UPLOAD_API } from "../utils/constants";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

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
    marginBottom: "0.7rem"
  },
  searchInput: {
    padding: "5px 10px",
    fontSize: "0.9rem",
    borderRadius: "5px",
    outline: "none",
    border: "1px solid black",
    boxSizing: "border-box"
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

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

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
        const file = hit?.resource;
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

  const fetchFiles = async () => {
    await callGraphApi("RECENT");
  }

  useEffect(() => {
    fetchFiles();
  }, [])

  return (
    <div className={styles.root}>
      {/* common */}
      {/* search, browse */}
      <div className={styles.searchDiv}>
        <input
          type="text"
          name="searchQuery"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className={styles.searchInput} />
        <button
          onClick={searchPpt}
          className={styles.searchButton}>Search</button>
      </div>
      {/* browse */}
      <FileInput
        setFile={setFile}
        generatePPTDetails={generatePPTDetails} />
      {error && <p className={styles.block}>{error}</p>}

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
            setShowSlides={setShowSlides}
            sourceSlideIds={sourceSlideIds}
            formatting={formatting}
            setFormatting={setFormatting} />
        )}
    </div>
  );
};

export default App;
