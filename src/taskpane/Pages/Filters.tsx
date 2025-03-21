import { Button, makeStyles } from "@fluentui/react-components";
import { DismissFilled } from "@fluentui/react-icons";
import React, { Dispatch, SetStateAction, useState } from "react";

interface IFilterPage {
    filterOptions: any;
    userFilter: any;
    setUserFilter: any;
    setFilterPage: Dispatch<SetStateAction<boolean>>;
    handleFilter: ({ }) => void;
    setUiResults: React.Dispatch<React.SetStateAction<any[]>>;
}

const useStyles = makeStyles({
    filterItem: {
        display: "flex",
        marginBottom: "0.3rem"
    },
    filterLabel: {
        fontSize: "1rem",
        marginBottom: "0.5rem",
        display: "block"
    },
    btnsDiv: {
        display: "flex",
        justifyContent: "space-evenly",
        marginTop: "1.5rem"
    },
    fil__head: {
        display: "flex",
        alignItems: "center",
        marginBottom: "1.2rem"
    }
})

const Filters = (props: IFilterPage) => {
    const { filterOptions, userFilter, setUserFilter, setFilterPage, handleFilter } = props;
    const styles = useStyles();
    const [localState, setLocalState] = useState(userFilter);

    const handleChange = (event, fieldName: string) => {
        const value = event.target.value;

        const fieldVals = localState?.[fieldName] || [];
        // Check if the value is already selected
        if (localState?.[fieldName]?.includes(value)) {
            setLocalState((prev) => {
                return {
                    ...prev,
                    [fieldName]: fieldVals?.filter(tag => tag !== value)
                }
            });
        } else {
            setLocalState((prev) => {
                return {
                    ...prev,
                    [fieldName]: prev?.[fieldName] ? [...prev?.[fieldName], value] : [value]
                }
            });
        }
    };

    const handleDisabled = (): boolean => {
        let disabled = false;
        for (const key in localState) {
            if (!userFilter || !userFilter?.[key]) {
                disabled = true;
                break;
            } else {
                const currValues = localState?.[key];
                const parentValues = userFilter?.[key];
                if (currValues?.length !== parentValues?.length) {
                    disabled = false;
                    break;
                }
            }
            // else {
            //     const val = parentValues?.find((val) => currValues?.some((cv) => cv === val));
            //     disabled = !Boolean(val);
            // }
        }
        return disabled;
    }

    const handleClear = () => {
        if (!userFilter) return;
        setUserFilter({});
        setLocalState({});
        handleFilter({ type: "FILTER" });
    }

    const applyFilters = () => {
        if (!userFilter) return;
        setUserFilter(localState);
        setFilterPage(false);
        handleFilter({ filterKey: "Tag", filterValue: localState?.["tag"] });
    }

    return (
        <div>
            <div className={styles.fil__head}>
                <h3 style={{ margin: "0", marginRight: "0.5rem" }}>Filters</h3>
                <DismissFilled fontSize="20px" style={{ cursor: "pointer" }} onClick={() => setFilterPage(false)} />
            </div>
            {filterOptions?.map((filter) => {
                return (
                    <div key={filter?.name}>
                        <label className={styles.filterLabel}>{filter?.name}</label>
                        <div>
                            {filter?.choice?.choices?.map((choice) => {
                                return (
                                    <div key={choice} className={styles.filterItem}>
                                        <label>
                                            <input
                                                type="checkbox"
                                                name={filter?.name?.toLowerCase()}
                                                value={choice}
                                                checked={
                                                    localState?.[filter?.name?.toLowerCase()]
                                                        ? localState?.[filter?.name?.toLowerCase()].includes(choice)
                                                        : false}
                                                onChange={(e) => handleChange(e, filter?.name?.toLowerCase())} />
                                            {choice}
                                        </label>
                                    </div>
                                )
                            })}
                        </div>
                    </div>
                )
            })}

            <div className={styles.btnsDiv}>
                <Button
                    shape="circular"
                    // disabled={handleDisabled()} 
                    onClick={handleClear}>
                    Clear
                </Button>
                <Button
                    appearance="primary"
                    shape="circular"
                    // disabled={handleDisabled()}
                    onClick={applyFilters}>
                    Apply filters
                </Button>
            </div>
        </div>
    );
};

export default Filters;
