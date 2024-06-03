// Converts a CAML case string to upper case first letter and spaces
export function convertCAML(value: string = "") {
    if (value.length == 0) { return value; }

    // Convert the string
    return value.replace(/(?=[A-Z])/g, " ").replace(/^./, value[0].toUpperCase());
}