import React from "react";
import { CustomDriveItemResponse } from "../Types";
import { Button, Spinner } from "@fluentui/react-components";

interface IPpt {
    searchResults: any[];
    generatePPTDetails: (x: CustomDriveItemResponse) => void;
    isSearchClicked: boolean;
    loading: boolean;
    handleOpenInPpt: (index: number) => void;
}

const Ppt = (props: IPpt) => {
    const { searchResults, generatePPTDetails, isSearchClicked, loading, handleOpenInPpt } = props;

    const handleClick = (result: CustomDriveItemResponse) => {
        generatePPTDetails(result);
    }

    if (loading) {
        return <p style={{
            width: "100%",
            height: "75vh",
            display: "flex",
            justifyContent: "center",
            alignItems: "center"
        }}><Spinner /></p>
    }

    return (
        <div style={{ marginTop: "1rem" }}>
            {isSearchClicked && <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <h2>Results</h2>
            </div>}
            {searchResults?.length === 0 && !loading ? (
                <p>No results found</p>
            ) : (
                <>
                    {searchResults?.map((result, idx) => {
                        return (
                            <div key={result?.id} style={{ width: "100%", marginBottom: "0.5rem" }}>
                                <p style={{
                                    fontWeight: "500",
                                    marginBlock: "8px"
                                }}>{result?.fields?.FileLeafRef || result?.name}</p>
                                <img
                                    src={result?.thumbnail}
                                    alt=""
                                    onClick={() => handleClick(result)}
                                    style={{
                                        aspectRatio: "3 / 4",
                                        width: "100%",
                                        maxHeight: "150px",
                                        borderRadius: "8px",
                                        cursor: "pointer"
                                    }} />
                                <div style={{ width: "100%", display: "flex", justifyContent: "flex-end" }}>
                                    <Button size="small" onClick={() => handleOpenInPpt(idx)}>Open in PPT</Button>
                                </div>
                            </div>
                        )
                    })}
                </>
            )}
        </div>
    );
};

export default Ppt;
