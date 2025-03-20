export const getFilteredData = (datas, key, value, removing = false) => {
    if (value?.length === 0) return datas;
    if (removing) {
        return datas?.filter((data) => {
            const val = value?.some((val) => data?.fields[key]?.includes(val));
            return val;
        });
    } else {
        return datas?.filter((data) => {
            const val = value?.some((val) => !data?.fields[key]?.includes(val));
            return val;
        });
    }
}