export const getFilteredData = (datas: any, key: string, value: string[]) => {
    if (value?.length === 0) return datas;
    return datas?.filter((data) => {
        const val = value?.some((val) => {
            return data?.fields?.[key]?.includes(val)
        });
        return val;
    });
}