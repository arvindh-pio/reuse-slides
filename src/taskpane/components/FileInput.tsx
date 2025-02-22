import React, { useRef, useState } from "react";
import { API_BASE_URL, UPLOAD_API } from "../utils/constants";
import { makeStyles } from "@fluentui/react-components";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { ISlide } from "./App";

interface IFileInput {
    setFile: React.Dispatch<React.SetStateAction<File | null>>;
    setPreviews: React.Dispatch<React.SetStateAction<string[] | null>>;
    setSourceSlideIds: React.Dispatch<React.SetStateAction<ISlide[] | null>>;
    setBase64: React.Dispatch<React.SetStateAction<string | null>>;
    setFormatting: React.Dispatch<React.SetStateAction<Boolean | null>>;
    formatting: boolean;
    sourceSlidesLength: number;
}

const useStyles = makeStyles({
    inputContainer: {
        display: "flex",
        justifyContent: "center",
        flexDirection: "column",
        textAlign: "center"
    },
    inputFile: {
        display: 'none'
    },
    block: {
        display: "block",
        margin: "0"
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
    formatDiv: {
        display: "flex",
        gap: "4px",
        margin: "10px auto 14px"
    }
})

const FileInput = (props: IFileInput) => {
    const styles = useStyles();
    const { setFile, setPreviews, setSourceSlideIds, setBase64, setFormatting, formatting, sourceSlidesLength } = props;
    const fileInputRef = useRef(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const handleBrowseClick = () => {
        fileInputRef.current.click();
    };

    const handleFileChange = async (event) => {
        const file = event.target.files[0];
        if (!file) {
            return;
        }

        setFile(file);
        generatePreviews(file);
        const slideIds = await extractSlideIds(file);
        setSourceSlideIds(slideIds);
        const base64 = await getBase64(file);
        setBase64(base64 as string);
    };

    const generatePreviews = async (file: File) => {
        try {
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
            console.log("Error -> ", error);
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

    return (
        <div className={styles.inputContainer}>
            <button type="button" onClick={handleBrowseClick} className={styles.browseButton}>
                Browse
            </button>
            {error && <p className={styles.block}>{error}</p>}
            <input
                type="file"
                ref={fileInputRef}
                accept=".pptx"
                onChange={handleFileChange}
                className={styles.inputFile}
            />
            {loading && <p className={styles.block}>Loading...</p>}

            {/* source formatting */}
            <div className={styles.formatDiv}>
                <input 
                    type="checkbox" 
                    name="formatting"
                    id="formatting"
                    checked={formatting}
                    disabled={!sourceSlidesLength}
                    onChange={() => setFormatting((prev) => !prev)} />
                <label htmlFor="formatting">Keep source formatting</label>
            </div>
        </div>
    )
};

export default FileInput;
