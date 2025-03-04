export const getFilteredData = (datas, key, value) => {
    return datas?.filter((data) => data?.fields[key].includes(value));
}