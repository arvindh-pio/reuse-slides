export const getFilteredData = (datas, key, value) => {
    if (value?.length === 0) return datas;
    return datas?.filter((data) => {
        const val = value?.some((val) => data?.fields[key]?.includes(val));
        return val;
    });
}