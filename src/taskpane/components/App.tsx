import React, { useEffect, useState } from "react";
import { Button, Input, makeStyles, Spinner } from "@fluentui/react-components";
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
import { FilterAddFilled, FilterRegular } from "@fluentui/react-icons";
import useInitial from "../hooks/useInitial";
import useFiles from "../hooks/useFiles";
import Filters from "../Pages/Filters";

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
    // padding: "5px 10px",
    // fontSize: "0.9rem",
    // borderRadius: "5px",
    // outline: "none",
    // border: "1px solid black",
    // boxSizing: "border-box",
    width: "100%",
  },
  actionBtns: {
    display: "flex",
    width: "100%",
    justifyContent: "space-evenly",
    marginBottom: "0.7rem",
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
  },
  filter: {
    fontSize: "1.5rem",
    cursor: "pointer",
    padding: "3px",
    borderRadius: "5px",
    ":hover": {
      backgroundColor: "rgba(0,0,0,0.1)"
    }
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
  const [filterPage, setFilterPage] = useState(false);
  const [filterOptions, setFilterOptions] = useState([]);
  const [userFilter, setUserFilter] = useState({});

  // 2
  const { fetchPPTFiles } = useInitial();
  const [filteredResults, setFilteredResults] = useState([]);

  // backend call for generate previews
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

  const searchForKeywordInLibraryDocs = async () => {
    setError(null);
    setSearchClicked(true);
    setLoading(true);
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
      // const modData = filteredResults?.length > 0 ? filteredResults : searchResults?.length > 0 ? searchResults : recentResults;
      const filteredFiles = recentResults?.filter((file: any) => {
        const recentFiles = pptFiles?.some((result: any) => result?.name === file?.fields?.FileLeafRef);
        return recentFiles;
      });
      setSearchResults(filteredFiles);
    } catch (error) {
    } finally {
      setLoading(false);
    }
  }

  const handleReset = () => {
    setSearchQuery("");
    setSearchResults([]);
    setSearchClicked(false);
    setError("");
  }

  const handleFilter = async (key: string, val: string) => {
    // const modData = filteredResults?.length > 0 ? filteredResults : searchResults?.length > 0 ? searchResults : recentResults;
    const datas = getFilteredData(recentResults, key, val, true);
    setFilteredResults(datas);
  }

  // testing purpose for filter api
  // const test = async () => {
  //   const token = getToken();
  //   const library = await getLibraryId(site.id);

  //   try {
  //     const response = await fetch(
  //       `https://graph.microsoft.com/v1.0/sites/${site.id}/drives/b!A8xRw8PhhEipKKGLOg8jIdLYbQIfU9pOn4DXZs7wMSE2CHN5BdKHRr8UO3rRbCpK/root/search(q='presentation')?$filter=fields/Tag eq 'Tech'`, {
  //       headers: {
  //         Authorization: `Bearer ${token}`,
  //         "Content-Type": "application/json"
  //       },
  //       method: "GET"
  //     });
  //     const colsData = await response.json();
  //     console.log(colsData, "^^^^^^");
  //   } catch (error) {
  //     console.log("filters error -> ", error);
  //   }
  // }

  const fetchConfig = async () => {
    setLoading(true);
    try {
      const response = await fetch(API_BASE_URL + "/config", {
        method: "GET",
        headers: {
          "Content-Type": "application/json"
        },
      });
      const data = await response.json();
      setSiteName(data?.siteName);
      setLibraryName(data?.libraryName);

      const { files, drive, site, filterConfigs } = await fetchPPTFiles(data);
      setRecentResults(files);
      setDrive(drive);
      setSite(site);
      setFilterOptions(filterConfigs);
    } catch (error) {
      console.log("Error in fetching config ", error);
    }
    setLoading(false);
  }

  const checkValue = () => {
    let disabled = true;
    for (const key in userFilter) {
      const values = userFilter?.[key];
      if (values?.length > 0) {
        disabled = false;
        break;
      }
    }
    return disabled;
  }

  useEffect(() => {
    if (!showSlides) {
      setSearchQuery("");
    }
  }, [showSlides])

  useEffect(() => {
    fetchConfig();
  }, [])

  if (loading || !siteName || !libraryName) return <p className={styles.loading}><Spinner /></p>

  return (
    <div className={styles.root}>
      <div className={styles.searchDiv}>
        <Input
          type="text"
          name="searchQuery"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className={styles.searchInput} />
        <div className={styles.actionBtns}>
          <Button shape="circular" appearance="primary" onClick={searchForKeywordInLibraryDocs} disabled={!searchQuery}>Search</Button>
          <Button shape="circular" onClick={handleReset} disabled={!isSearchClicked}>Reset</Button>
          <div onClick={() => setFilterPage(true)}>
            {checkValue()
              ? <FilterRegular className={styles.filter} />
              : <FilterAddFilled className={styles.filter} />}
          </div>
        </div>
      </div>
      {/* uncomment to add browse functionality */}
      {/* <FileInput
            setFile={setFile}
            generatePPTDetails={generatePPTDetails} />
          {error && <p className={styles.block}>{error}</p>} */}

      {filterPage ? (
        <Filters 
          filterOptions={filterOptions} 
          userFilter={userFilter} 
          setUserFilter={setUserFilter} 
          setFilterPage={setFilterPage}
          handleFilter={handleFilter}
          setFilteredResults={setFilteredResults} />
      ) : !showSlides
        ? (
          <>
            <Ppt
              searchResults={
                filteredResults?.length > 0
                  ? filteredResults
                  : (searchResults?.length > 0 || isSearchClicked) ? searchResults : recentResults}
              generatePPTDetails={generatePPTDetails}
              isSearchClicked={isSearchClicked}
              loading={loading} />
          </>
        )
        : loading ? (
          <Spinner />
        ) : (
          <Slides
            base64={base64}
            previews={previews}
            sourceSlideIds={sourceSlideIds}
            formatting={formatting}
            setFormatting={setFormatting}
            handleBack={handleBack} />
        )}
    </div>
  );
};

export default App;
