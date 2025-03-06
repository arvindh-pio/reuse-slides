import React, { useState } from "react";
import { getToken } from "../utils/utils";

const FilterDropdown = (props: any) => {
    const { customObject, site, library, handleFilter } = props;
    const [selectedOption, setSelectedOption] = useState("All");

    // Array of options to be displayed in the dropdown
    const options = customObject?.choice?.choices;

    // Handle the change in selection
    const handleChange = async (event) => {
        const val = event.target.value;
        setSelectedOption(val);
        handleFilter(customObject?.displayName, val);

        // const url = 
        //     `https://graph.microsoft.com/v1.0/sites/${site.id}/drives/${library.id}/root/search(q='presentation')?$filter=fields/Tag eq '${val}'`;

        // try {
        //     const token = getToken();
        //     const response = await fetch(url, {
        //         headers: {
        //             Authorization: `Bearer ${token}`,
        //             "Content-Type": "application/json"
        //         },
        //         method: "GET"
        //     });
        //     const data = await response.json();
        //     console.log("*******data ", data);
        // } catch (error) {
        //     console.log("err ", error);
        // }
    };

    return (
        <>
            {options?.map((option) => {
                return (
                    <>
                        <input type="checkbox" id="tag" name="tag" value={option} onChange={handleChange} />
                        <label htmlFor="tag">{option}</label>
                    </>
                )
            })}
            <br />
        </>
    );
};

export default FilterDropdown;
