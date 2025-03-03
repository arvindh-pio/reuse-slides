import { makeStyles } from "@fluentui/react-components";
import React from "react";

interface IConfig {
    siteName: string;
    setSiteName: React.Dispatch<React.SetStateAction<string>>;
    libraryName: string;
    setLibraryName: React.Dispatch<React.SetStateAction<string>>;
    getRecentFilesFromLibrary: () => void;
}

const useStyles = makeStyles({
    configDiv: {
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        height: "73vh",
    },
    input: {
        padding: "0.5rem",
        marginBottom: "0.5rem",
        fontSize: "0.8rem",
        fontFamily: "sans-serif"
    },
    button: {
        cursor: "pointer",
        padding: "0.5rem",
        marginTop: "0.7rem",
        fontFamily: "sans-serif",
        backgroundColor: "crimson",
        border: "0",
        borderRadius: "6px",
        color: "white",
        ":hover": {
            opacity: "0.8"
        }
    }
})

const Config = (props: IConfig) => {
    const { siteName, setSiteName, libraryName, setLibraryName, getRecentFilesFromLibrary } = props;
    const styles = useStyles();

    return (
        <div className={styles.configDiv}>
            <input
                className={styles.input}
                name="siteName"
                placeholder="Enter site name"
                value={siteName}
                onChange={(e) => setSiteName(e.target.value)} />
            <input
                className={styles.input}
                name="libraryName"
                placeholder="Enter document library name"
                value={libraryName}
                onChange={(e) => setLibraryName(e.target.value)} />
            <button className={styles.button} onClick={getRecentFilesFromLibrary}>Go</button>
        </div>
    );
};

export default Config;
