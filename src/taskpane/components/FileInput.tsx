import React, { useRef } from "react";
import { makeStyles } from "@fluentui/react-components";

interface IFileInput {
    setFile: React.Dispatch<React.SetStateAction<File | null>>;
    generatePPTDetails: (x: File) => void;
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
    }
})

const FileInput = (props: IFileInput) => {
    const styles = useStyles();
    const { setFile, generatePPTDetails } = props;
    const fileInputRef = useRef(null);

    const handleBrowseClick = () => {
        fileInputRef.current.click();
    };

    const handleFileChange = async (event) => {
        const file = event.target.files[0];
        if (!file) {
            return;
        }

        setFile(file);
        await generatePPTDetails(file);
    };

    return (
        <div className={styles.inputContainer}>
            <button type="button" onClick={handleBrowseClick} className={styles.browseButton}>
                Browse
            </button>
            <input
                type="file"
                ref={fileInputRef}
                accept=".pptx"
                onChange={handleFileChange}
                className={styles.inputFile}
            />
        </div>
    )
};

export default FileInput;
