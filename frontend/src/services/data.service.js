import axios from "axios";

const API_URL = "http://10.0.0.236:5000/api";

export const getData = async () => {
  const res = await axios.get(`${API_URL}/data`);
  return res.data;
};
