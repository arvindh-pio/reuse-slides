import React from "react";
import { CustomDriveItemResponse } from "../Types";

interface IPpt {
    searchResults: CustomDriveItemResponse[];
}

const Ppt = (props: IPpt) => {
    const { searchResults } = props;

    return (
        <div>
            <h2 style={{ margin: "0 0 1rem 0" }}>Results</h2>
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

export default Ppt;
