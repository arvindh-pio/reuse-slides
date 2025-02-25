import React, { useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import FileInput from "./FileInput";
import Previews from "./Previews";
import Ppt from "../Pages/Ppt";
import Slides from "../Pages/Slides";
import { CustomDriveItemResponse, DriveItemResponse } from "../Types";

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
  const [searchResults, setSearchResults] = useState<CustomDriveItemResponse[]>([]);

  const searchPpt = async () => {
    const token = localStorage.getItem("token");
    if (!searchQuery) return;

    const response = await
      fetch(`https://graph.microsoft.com/v1.0/me/drive/root/search(q=\'${searchQuery}\')`,
        {
          method: "GET",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          }
        }
      );

    const data: DriveItemResponse = await response.json();
    const pptFiles = await Promise.all(
      data.value.map(async (file) => {
        const extArr = file.name.split(".") || [];
        const ext = extArr[extArr.length - 1];
        console.log("file ", file, extArr, ext);

        // Check if it's a PowerPoint file (ppt or pptx)
        if (ext === "ppt" || ext === "pptx") {
          // Fetch thumbnail for the file
          const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/thumbnails`, {
            method: "GET",
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json"
            }
          });
          const data = await response.json();
          const mediumImage = data?.value?.[0]?.large?.url;

          // Return the file with thumbnail information
          return {
            thumbnail: mediumImage,
            ...file
          };
        }
        return null; // Return null for non-PPT files
      })
    );

    // Filter out null values (non-PPT files)
    const pptResults = pptFiles.filter(file => file !== null);

    // Set the search results with PPT files and their thumbnails
    setSearchResults(pptResults);
  }

  return (
    <div className={styles.root}>
      {/* common */}
      {/* search */}
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
        setPreviews={setPreviews}
        setSourceSlideIds={setSourceSlideIds}
        setBase64={setBase64}
        setFormatting={setFormatting}
        sourceSlidesLength={sourceSlideIds?.length}
        formatting={formatting} />

      {/* ppt */}
      {searchResults?.length > 0 && <Ppt searchResults={searchResults} />}

      {/* slides */}
      <Slides />

      {/* <FileInput
        setFile={setFile}
        setPreviews={setPreviews}
        setSourceSlideIds={setSourceSlideIds}
        setBase64={setBase64}
        setFormatting={setFormatting}
        sourceSlidesLength={sourceSlideIds?.length}
        formatting={formatting} />
      <Previews base64={base64} previews={previews} sourceSlideIds={sourceSlideIds} formatting={formatting} /> */}
    </div>
  );
};

export default App;
