import axios from "axios";

// const API_URL = "http://10.0.0.236:5000/api";
const API_URL = "http://localhost:5000/api";

export const getData = async () => {
  const res = await axios.get(`${API_URL}/data`);
  return res.data;
};

export const resetData = async () => {
  const res = await axios.post(`${API_URL}/excel/reset`);
  return res.data;
}

export const importExcel = async (data) => {
  const res = await axios.post(`${API_URL}/excel/preview`, data);
  return res.data;
}