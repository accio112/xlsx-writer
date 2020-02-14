interface Dimension{
    col: Number;
    row: Number;
}
interface Merge {
    start: Number;
    end?: Number;
}
interface Style {
    name?: String;
    size?: String;
    bold?: String;
    underline?: String;
    color?: String;
    bgColor?: String;
    fgColor?: String;
    border?: String;
}

interface Data{
    name? :String;
    topLeft?:  Dimension;
    bottomRight?: Dimension;
    mergeCells?: Merge;
    style?: Style;
}
interface Fields{
   data: Data; 
}

interface table{
    headers: Fields;
    rowsData: Data[][];
}
export interface XlsxData {
    worksheetName: String;
    image?: Fields;
    title?: Fields;
    tables?: table[];
}
