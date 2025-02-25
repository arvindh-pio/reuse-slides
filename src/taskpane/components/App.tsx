import React, { useEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import FileInput from "./FileInput";
import Previews from "./Previews";
import { msalInstance } from "../../configs/authConfig";

export interface ISlide {
  index: number;
  slideId: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "1rem"
  },
});

const App: React.FC = () => {
  // styles
  const styles = useStyles();
  const [isUserLoggedIn, setIsUserLoggedIn] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");

  // states
  const [file, setFile] = useState<File | null>(null);
  const [previews, setPreviews] = useState<string[]>([]);
  const [sourceSlideIds, setSourceSlideIds] = useState<ISlide[]>([]);
  const [base64, setBase64] = useState<string | null>(null);
  const [formatting, setFormatting] = useState(true);
  const [searchResults, setSearchResults] = useState([]);

  useEffect(() => {
    handleLogin();
  }, []);

  const handleLogin = async () => {
    try {
      await msalInstance.initialize();
      const loginResponse = await msalInstance.loginPopup({
        scopes: ["User.read", "Files.Read", "Files.Read.All", "Sites.Read.All"]
      });
      console.log("login res -> ", loginResponse);
    } catch (error) {
      console.log("Login failed -> ", error);
    }
  }

  const getAccessToken = async () => {
    try {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length === 0) {
        throw new Error("No account logged in");
      }

      const response = await msalInstance.acquireTokenSilent({
        scopes: ["Files.Read", "Files.Read.All"],
        account: accounts[0],
      })

      return response.accessToken;
    } catch (error) {
      console.log("Error getting token -> ", error);
      return null;
    }
  }

  const searchPpt = async () => {
    const token = await getAccessToken();
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

    const data = await response.json();
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
      <div style={{
        display: "flex",
        gap: "5px",
        alignItems: "center",
        width: "100%",
        boxSizing: "border-box",
        marginBottom: "0.7rem"
      }}>
        <input
          type="text"
          name="searchQuery"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          style={{
            padding: "5px 10px",
            fontSize: "0.9rem",
            borderRadius: "5px",
            outline: "none",
            border: "1px solid black",
            boxSizing: "border-box"
          }} />
        <button
          onClick={searchPpt}
          style={{
            background: "none",
            border: "1px solid crimson",
            padding: "5px 10px",
            fontFamily: "sans-serif",
            borderRadius: "4px",
            cursor: "pointer"            
          }}>Search</button>
      </div>
      <FileInput
        setFile={setFile}
        setPreviews={setPreviews}
        setSourceSlideIds={setSourceSlideIds}
        setBase64={setBase64}
        setFormatting={setFormatting}
        sourceSlidesLength={sourceSlideIds?.length}
        formatting={formatting} />
      <Previews base64={base64} previews={previews} sourceSlideIds={sourceSlideIds} formatting={formatting} />
      {searchResults?.length > 0 && <h2 style={{ margin: "0 0 1rem 0" }}>Results</h2>}
      {searchResults?.map((result) => {
        return (
          <div style={{ width: "100%", marginBottom: "0.5rem" }}>
            <p style={{
              fontWeight: "500",
              marginBlock: "5px"
            }}>{result?.name}</p>
            <img src={result?.thumbnail} alt="" style={{
              aspectRatio: "3 / 4",
              width: "100%",
              maxHeight: "150px",
              borderRadius: "8px"
            }} />
          </div>
        )
      })}
    </div>
  );
};

export default App;
