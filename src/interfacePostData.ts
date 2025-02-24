type Kvs = {
  keys: { [key: string]: string };
  values: { [key: string]: string };
}[];
type Dict = {
  destination: string;
  data: Kvs;
};
