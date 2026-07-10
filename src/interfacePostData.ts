type Kvs = {
  keys: {[key: string]: string};
  values: {[key: string]: string};
}[];
// eslint-disable-next-line @typescript-eslint/no-unused-vars
type Dict = {
  destination: string;
  data: Kvs;
};
