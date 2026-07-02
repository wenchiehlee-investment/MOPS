# MOPS Downloader System

[![Version](https://img.shields.io/badge/Version-2.0.0-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Architecture](https://img.shields.io/badge/Architecture-Clean_Pipeline-orange)](https://github.com/your-repo/mops-downloader)

A Python-based tool for automatically downloading quarterly financial reports from Taiwan's Market Observation Post System (MOPS). Designed to handle real-world variations in report availability with intelligent fallback mechanisms.

<!-- BEGIN_STATUS -->

## Current MOPS PDFs

> **Source**: `mops_matrix_latest.csv`

| 代號 | 名稱 | 2026 Q2 | 2026 Q1 | 2025 Q4 | 2025 Q3 | 2025 Q2 | 2025 Q1 | 2024 Q4 | 2024 Q3 | 2024 Q2 | 2024 Q1 | 2023 Q4 | 2023 Q3 | 2023 Q2 | 2023 Q1 | 2020 Q4 | 2020 Q3 | 2020 Q2 | 2020 Q1 | process_timestamp |
|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|------|
| 1587 | 吉茂 | - | [AI1](downloads/1587/202601_1587_AI1.pdf) | [AI1](downloads/1587/202504_1587_AI1.pdf) / [AI3](downloads/1587/202504_1587_AI3.pdf) | [AI1](downloads/1587/202503_1587_AI1.pdf) | [AI1](downloads/1587/202502_1587_AI1.pdf) | [AI1](downloads/1587/202501_1587_AI1.pdf) | [AI1](downloads/1587/202404_1587_AI1.pdf) / [AI3](downloads/1587/202404_1587_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2301 | 光寶科 | - | [AI1](downloads/2301/202601_2301_AI1.pdf) | [AI1](downloads/2301/202504_2301_AI1.pdf) / [AI3](downloads/2301/202504_2301_AI3.pdf) | [AI1](downloads/2301/202503_2301_AI1.pdf) | [AI1](downloads/2301/202502_2301_AI1.pdf) | [AI1](downloads/2301/202501_2301_AI1.pdf) | [AI1](downloads/2301/202404_2301_AI1.pdf) / [AI3](downloads/2301/202404_2301_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2303 | 聯電 | - | [AI1](downloads/2303/202601_2303_AI1.pdf) | [AI1](downloads/2303/202504_2303_AI1.pdf) / [AI3](downloads/2303/202504_2303_AI3.pdf) | [AI1](downloads/2303/202503_2303_AI1.pdf) | [AI1](downloads/2303/202502_2303_AI1.pdf) | [AI1](downloads/2303/202501_2303_AI1.pdf) | [AI1](downloads/2303/202404_2303_AI1.pdf) / [AI3](downloads/2303/202404_2303_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2308 | 台達電 | - | [AI1](downloads/2308/202601_2308_AI1.pdf) | [AI1](downloads/2308/202504_2308_AI1.pdf) / [AI3](downloads/2308/202504_2308_AI3.pdf) | [AI1](downloads/2308/202503_2308_AI1.pdf) | [AI1](downloads/2308/202502_2308_AI1.pdf) | [AI1](downloads/2308/202501_2308_AI1.pdf) | [AI1](downloads/2308/202404_2308_AI1.pdf) / [AI3](downloads/2308/202404_2308_AI3.pdf) | - | - | - | [AI1](downloads/2308/202304_2308_AI1.pdf) / [AI3](downloads/2308/202304_2308_AI3.pdf) | [AI1](downloads/2308/202303_2308_AI1.pdf) | [AI1](downloads/2308/202302_2308_AI1.pdf) | [AI1](downloads/2308/202301_2308_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2317 | 鴻海 | - | [AI1](downloads/2317/202601_2317_AI1.pdf) | [AI1](downloads/2317/202504_2317_AI1.pdf) / [AI3](downloads/2317/202504_2317_AI3.pdf) | [AI1](downloads/2317/202503_2317_AI1.pdf) | [AI1](downloads/2317/202502_2317_AI1.pdf) | [AI1](downloads/2317/202501_2317_AI1.pdf) | [AI1](downloads/2317/202404_2317_AI1.pdf) / [AI3](downloads/2317/202404_2317_AI3.pdf) | - | - | - | [AI1](downloads/2317/202304_2317_AI1.pdf) / [AI3](downloads/2317/202304_2317_AI3.pdf) | [AI1](downloads/2317/202303_2317_AI1.pdf) | [AI1](downloads/2317/202302_2317_AI1.pdf) | [AI1](downloads/2317/202301_2317_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2324 | 仁寶 | - | [AI1](downloads/2324/202601_2324_AI1.pdf) | [AI1](downloads/2324/202504_2324_AI1.pdf) / [AI3](downloads/2324/202504_2324_AI3.pdf) | [AI1](downloads/2324/202503_2324_AI1.pdf) | [AI1](downloads/2324/202502_2324_AI1.pdf) | [AI1](downloads/2324/202501_2324_AI1.pdf) | [AI1](downloads/2324/202404_2324_AI1.pdf) / [AI3](downloads/2324/202404_2324_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2330 | 台積電 | - | [AI1](downloads/2330/202601_2330_AI1.pdf) | [AI1](downloads/2330/202504_2330_AI1.pdf) / [AI3](downloads/2330/202504_2330_AI3.pdf) | [AI1](downloads/2330/202503_2330_AI1.pdf) | [AI1](downloads/2330/202502_2330_AI1.pdf) | [AI1](downloads/2330/202501_2330_AI1.pdf) | [AI1](downloads/2330/202404_2330_AI1.pdf) / [AI3](downloads/2330/202404_2330_AI3.pdf) | [AI1](downloads/2330/202403_2330_AI1.pdf) | [AI1](downloads/2330/202402_2330_AI1.pdf) | [AI1](downloads/2330/202401_2330_AI1.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2332 | 友訊 | - | [AI1](downloads/2332/202601_2332_AI1.pdf) | [AI1](downloads/2332/202504_2332_AI1.pdf) | [AI1](downloads/2332/202503_2332_AI1.pdf) | [AI1](downloads/2332/202502_2332_AI1.pdf) | [AI1](downloads/2332/202501_2332_AI1.pdf) | [AI1](downloads/2332/202404_2332_AI1.pdf) / [AI3](downloads/2332/202404_2332_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2345 | 智邦 | - | [AI1](downloads/2345/202601_2345_AI1.pdf) | [AI1](downloads/2345/202504_2345_AI1.pdf) / [AI3](downloads/2345/202504_2345_AI3.pdf) | [AI1](downloads/2345/202503_2345_AI1.pdf) | [AI1](downloads/2345/202502_2345_AI1.pdf) | - | [AI1](downloads/2345/202404_2345_AI1.pdf) / [AI3](downloads/2345/202404_2345_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2347 | 聯強 | - | [AI1](downloads/2347/202601_2347_AI1.pdf) | [AI1](downloads/2347/202504_2347_AI1.pdf) / [AI3](downloads/2347/202504_2347_AI3.pdf) | [AI1](downloads/2347/202503_2347_AI1.pdf) | [AI1](downloads/2347/202502_2347_AI1.pdf) | [AI1](downloads/2347/202501_2347_AI1.pdf) | [AI1](downloads/2347/202404_2347_AI1.pdf) / [AI3](downloads/2347/202404_2347_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2353 | 宏碁 | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2354 | 鴻準 | - | [AI1](downloads/2354/202601_2354_AI1.pdf) | [AI1](downloads/2354/202504_2354_AI1.pdf) / [AI3](downloads/2354/202504_2354_AI3.pdf) | [AI1](downloads/2354/202503_2354_AI1.pdf) | [AI1](downloads/2354/202502_2354_AI1.pdf) | [AI1](downloads/2354/202501_2354_AI1.pdf) | [AI1](downloads/2354/202404_2354_AI1.pdf) / [AI3](downloads/2354/202404_2354_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2356 | 英業達 | - | [AI1](downloads/2356/202601_2356_AI1.pdf) | [AI1](downloads/2356/202504_2356_AI1.pdf) | [AI1](downloads/2356/202503_2356_AI1.pdf) | [AI1](downloads/2356/202502_2356_AI1.pdf) | [AI1](downloads/2356/202501_2356_AI1.pdf) | [AI1](downloads/2356/202404_2356_AI1.pdf) / [AI3](downloads/2356/202404_2356_AI3.pdf) | [AI1](downloads/2356/202403_2356_AI1.pdf) | [AI1](downloads/2356/202402_2356_AI1.pdf) | [AI1](downloads/2356/202401_2356_AI1.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2357 | 華碩 | - | [AI1](downloads/2357/202601_2357_AI1.pdf) | [AI1](downloads/2357/202504_2357_AI1.pdf) / [AI3](downloads/2357/202504_2357_AI3.pdf) | [AI1](downloads/2357/202503_2357_AI1.pdf) | [AI1](downloads/2357/202502_2357_AI1.pdf) | [AI1](downloads/2357/202501_2357_AI1.pdf) | [AI1](downloads/2357/202404_2357_AI1.pdf) / [AI3](downloads/2357/202404_2357_AI3.pdf) | [AI1](downloads/2357/202403_2357_AI1.pdf) | - | - | - | - | - | - | [AI1](downloads/2357/202004_2357_AI1.pdf) / [AI3](downloads/2357/202004_2357_AI3.pdf) | [AI1](downloads/2357/202003_2357_AI1.pdf) | [AI1](downloads/2357/202002_2357_AI1.pdf) | [AI1](downloads/2357/202001_2357_AI1.pdf) | 2026-06-09T13:58:06.341789+08:00 |
| 2359 | 所羅門 | - | [AI1](downloads/2359/202601_2359_AI1.pdf) | [AI1](downloads/2359/202504_2359_AI1.pdf) / [AI3](downloads/2359/202504_2359_AI3.pdf) | [AI1](downloads/2359/202503_2359_AI1.pdf) | [AI1](downloads/2359/202502_2359_AI1.pdf) | - | [AI1](downloads/2359/202404_2359_AI1.pdf) / [AI3](downloads/2359/202404_2359_AI3.pdf) | [AI1](downloads/2359/202403_2359_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2376 | 技嘉 | - | [AI1](downloads/2376/202601_2376_AI1.pdf) | [AI1](downloads/2376/202504_2376_AI1.pdf) / [AI3](downloads/2376/202504_2376_AI3.pdf) | [AI1](downloads/2376/202503_2376_AI1.pdf) | [AI1](downloads/2376/202502_2376_AI1.pdf) | [AI1](downloads/2376/202501_2376_AI1.pdf) | [AI1](downloads/2376/202404_2376_AI1.pdf) / [AI3](downloads/2376/202404_2376_AI3.pdf) | [AI1](downloads/2376/202403_2376_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2377 | 微星 | - | [AI1](downloads/2377/202601_2377_AI1.pdf) | [AI1](downloads/2377/202504_2377_AI1.pdf) | [AI1](downloads/2377/202503_2377_AI1.pdf) | [AI1](downloads/2377/202502_2377_AI1.pdf) | [AI1](downloads/2377/202501_2377_AI1.pdf) | [AI1](downloads/2377/202404_2377_AI1.pdf) / [AI3](downloads/2377/202404_2377_AI3.pdf) | [AI1](downloads/2377/202403_2377_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2379 | 瑞昱 | - | [AI1](downloads/2379/202601_2379_AI1.pdf) | [AI1](downloads/2379/202504_2379_AI1.pdf) / [AI3](downloads/2379/202504_2379_AI3.pdf) | [AI1](downloads/2379/202503_2379_AI1.pdf) | [AI1](downloads/2379/202502_2379_AI1.pdf) | [AI1](downloads/2379/202501_2379_AI1.pdf) | [AI1](downloads/2379/202404_2379_AI1.pdf) / [AI3](downloads/2379/202404_2379_AI3.pdf) | [AI1](downloads/2379/202403_2379_AI1.pdf) | [AI1](downloads/2379/202402_2379_AI1.pdf) | [AI1](downloads/2379/202401_2379_AI1.pdf) | [AI1](downloads/2379/202304_2379_AI1.pdf) / [AI3](downloads/2379/202304_2379_AI3.pdf) | [AI1](downloads/2379/202303_2379_AI1.pdf) | [AI1](downloads/2379/202302_2379_AI1.pdf) | [AI1](downloads/2379/202301_2379_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2382 | 廣達 | - | [AI1](downloads/2382/202601_2382_AI1.pdf) | [AI1](downloads/2382/202504_2382_AI1.pdf) / [AI3](downloads/2382/202504_2382_AI3.pdf) | [AI1](downloads/2382/202503_2382_AI1.pdf) | [AI1](downloads/2382/202502_2382_AI1.pdf) | [AI1](downloads/2382/202501_2382_AI1.pdf) | [AI1](downloads/2382/202404_2382_AI1.pdf) / [AI3](downloads/2382/202404_2382_AI3.pdf) | [AI1](downloads/2382/202403_2382_AI1.pdf) | [AI1](downloads/2382/202402_2382_AI1.pdf) | [AI1](downloads/2382/202401_2382_AI1.pdf) | [AI1](downloads/2382/202304_2382_AI1.pdf) / [AI3](downloads/2382/202304_2382_AI3.pdf) | [AI1](downloads/2382/202303_2382_AI1.pdf) | [AI1](downloads/2382/202302_2382_AI1.pdf) | [AI1](downloads/2382/202301_2382_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2383 | 台光電 | - | [AI1](downloads/2383/202601_2383_AI1.pdf) | [AI1](downloads/2383/202504_2383_AI1.pdf) / [AI3](downloads/2383/202504_2383_AI3.pdf) | [AI1](downloads/2383/202503_2383_AI1.pdf) | [AI1](downloads/2383/202502_2383_AI1.pdf) | - | [AI1](downloads/2383/202404_2383_AI1.pdf) / [AI3](downloads/2383/202404_2383_AI3.pdf) | [AI1](downloads/2383/202403_2383_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2395 | 研華 | - | [AI1](downloads/2395/202601_2395_AI1.pdf) | [AI1](downloads/2395/202504_2395_AI1.pdf) / [AI3](downloads/2395/202504_2395_AI3.pdf) | [AI1](downloads/2395/202503_2395_AI1.pdf) | [AI1](downloads/2395/202502_2395_AI1.pdf) | [AI1](downloads/2395/202501_2395_AI1.pdf) | [AI1](downloads/2395/202404_2395_AI1.pdf) / [AI3](downloads/2395/202404_2395_AI3.pdf) | [AI1](downloads/2395/202403_2395_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2405 | 輔信 | - | [AI1](downloads/2405/202601_2405_AI1.pdf) | [AI1](downloads/2405/202504_2405_AI1.pdf) | [AI1](downloads/2405/202503_2405_AI1.pdf) | [AI1](downloads/2405/202502_2405_AI1.pdf) | - | [AI1](downloads/2405/202404_2405_AI1.pdf) | [AI1](downloads/2405/202403_2405_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2412 | 中華電 | - | [AI1](downloads/2412/202601_2412_AI1.pdf) | [AI1](downloads/2412/202504_2412_AI1.pdf) / [AI3](downloads/2412/202504_2412_AI3.pdf) | [AI1](downloads/2412/202503_2412_AI1.pdf) | [AI1](downloads/2412/202502_2412_AI1.pdf) | [AI1](downloads/2412/202501_2412_AI1.pdf) | [AI1](downloads/2412/202404_2412_AI1.pdf) / [AI3](downloads/2412/202404_2412_AI3.pdf) | [AI1](downloads/2412/202403_2412_AI1.pdf) | [AI1](downloads/2412/202402_2412_AI1.pdf) | [AI1](downloads/2412/202401_2412_AI1.pdf) | [AI1](downloads/2412/202304_2412_AI1.pdf) / [AI3](downloads/2412/202304_2412_AI3.pdf) | [AI1](downloads/2412/202303_2412_AI1.pdf) | [AI1](downloads/2412/202302_2412_AI1.pdf) | [AI1](downloads/2412/202301_2412_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2449 | 京元電子 | - | [AI1](downloads/2449/202601_2449_AI1.pdf) | [AI1](downloads/2449/202504_2449_AI1.pdf) / [AI3](downloads/2449/202504_2449_AI3.pdf) | [AI1](downloads/2449/202503_2449_AI1.pdf) | [AI1](downloads/2449/202502_2449_AI1.pdf) | [AI1](downloads/2449/202501_2449_AI1.pdf) | [AI1](downloads/2449/202404_2449_AI1.pdf) / [AI3](downloads/2449/202404_2449_AI3.pdf) | [AI1](downloads/2449/202403_2449_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2450 | 神腦 | - | [AI1](downloads/2450/202601_2450_AI1.pdf) | [AI1](downloads/2450/202504_2450_AI1.pdf) / [AI3](downloads/2450/202504_2450_AI3.pdf) | [AI1](downloads/2450/202503_2450_AI1.pdf) | [AI1](downloads/2450/202502_2450_AI1.pdf) | [AI1](downloads/2450/202501_2450_AI1.pdf) | [AI1](downloads/2450/202404_2450_AI1.pdf) / [AI3](downloads/2450/202404_2450_AI3.pdf) | [AI1](downloads/2450/202403_2450_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2451 | 創見 | - | [AI1](downloads/2451/202601_2451_AI1.pdf) | [AI1](downloads/2451/202504_2451_AI1.pdf) | [AI1](downloads/2451/202503_2451_AI1.pdf) | [AI1](downloads/2451/202502_2451_AI1.pdf) | [AI1](downloads/2451/202501_2451_AI1.pdf) | [AI1](downloads/2451/202404_2451_AI1.pdf) / [AI3](downloads/2451/202404_2451_AI3.pdf) | [AI1](downloads/2451/202403_2451_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2454 | 聯發科 | - | [AI1](downloads/2454/202601_2454_AI1.pdf) | [AI1](downloads/2454/202504_2454_AI1.pdf) | [AI1](downloads/2454/202503_2454_AI1.pdf) | [AI1](downloads/2454/202502_2454_AI1.pdf) | [AI1](downloads/2454/202501_2454_AI1.pdf) | [AI1](downloads/2454/202404_2454_AI1.pdf) / [AI3](downloads/2454/202404_2454_AI3.pdf) | [AI1](downloads/2454/202403_2454_AI1.pdf) | [AI1](downloads/2454/202402_2454_AI1.pdf) | [AI1](downloads/2454/202401_2454_AI1.pdf) | [AI1](downloads/2454/202304_2454_AI1.pdf) / [AI3](downloads/2454/202304_2454_AI3.pdf) | [AI1](downloads/2454/202303_2454_AI1.pdf) | [AI1](downloads/2454/202302_2454_AI1.pdf) | [AI1](downloads/2454/202301_2454_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2458 | 義隆 | - | [AI1](downloads/2458/202601_2458_AI1.pdf) | [AI1](downloads/2458/202504_2458_AI1.pdf) | [AI1](downloads/2458/202503_2458_AI1.pdf) | [AI1](downloads/2458/202502_2458_AI1.pdf) | [AI1](downloads/2458/202501_2458_AI1.pdf) | [AI1](downloads/2458/202404_2458_AI1.pdf) / [AI3](downloads/2458/202404_2458_AI3.pdf) | [AI1](downloads/2458/202403_2458_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2474 | 可成 | - | [AI1](downloads/2474/202601_2474_AI1.pdf) | [AI3](downloads/2474/202504_2474_AI3.pdf) | [AI1](downloads/2474/202503_2474_AI1.pdf) | [AI1](downloads/2474/202502_2474_AI1.pdf) | [AI1](downloads/2474/202501_2474_AI1.pdf) | [AI1](downloads/2474/202404_2474_AI1.pdf) / [AI3](downloads/2474/202404_2474_AI3.pdf) | [AI1](downloads/2474/202403_2474_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2480 | 敦陽科 | - | [AI1](downloads/2480/202601_2480_AI1.pdf) | [AI1](downloads/2480/202504_2480_AI1.pdf) / [AI3](downloads/2480/202504_2480_AI3.pdf) | [AI1](downloads/2480/202503_2480_AI1.pdf) | [AI1](downloads/2480/202502_2480_AI1.pdf) | [AI1](downloads/2480/202501_2480_AI1.pdf) | [AI1](downloads/2480/202404_2480_AI1.pdf) / [AI3](downloads/2480/202404_2480_AI3.pdf) | [AI1](downloads/2480/202403_2480_AI1.pdf) | [AI1](downloads/2480/202402_2480_AI1.pdf) | [AI1](downloads/2480/202401_2480_AI1.pdf) | [AI1](downloads/2480/202304_2480_AI1.pdf) / [AI3](downloads/2480/202304_2480_AI3.pdf) | [AI1](downloads/2480/202303_2480_AI1.pdf) | [AI1](downloads/2480/202302_2480_AI1.pdf) | [AI1](downloads/2480/202301_2480_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2603 | 長榮 | - | [AI1](downloads/2603/202601_2603_AI1.pdf) | [AI1](downloads/2603/202504_2603_AI1.pdf) / [AI3](downloads/2603/202504_2603_AI3.pdf) | [AI1](downloads/2603/202503_2603_AI1.pdf) | [AI1](downloads/2603/202502_2603_AI1.pdf) | [AI1](downloads/2603/202501_2603_AI1.pdf) | [AI1](downloads/2603/202404_2603_AI1.pdf) / [AI3](downloads/2603/202404_2603_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 2646 | 星宇航空 | - | [AI2](downloads/2646/202601_2646_AI2.pdf) | [AI2](downloads/2646/202504_2646_AI2.pdf) | [AI2](downloads/2646/202503_2646_AI2.pdf) | [AI2](downloads/2646/202502_2646_AI2.pdf) | [AI2](downloads/2646/202501_2646_AI2.pdf) | [AI2](downloads/2646/202404_2646_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3014 | 聯陽 | - | [AI1](downloads/3014/202601_3014_AI1.pdf) | [AI1](downloads/3014/202504_3014_AI1.pdf) / [AI3](downloads/3014/202504_3014_AI3.pdf) | [AI1](downloads/3014/202503_3014_AI1.pdf) | [AI1](downloads/3014/202502_3014_AI1.pdf) | [AI1](downloads/3014/202501_3014_AI1.pdf) | [AI1](downloads/3014/202404_3014_AI1.pdf) / [AI3](downloads/3014/202404_3014_AI3.pdf) | [AI1](downloads/3014/202403_3014_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3022 | 威強電 | - | [AI1](downloads/3022/202601_3022_AI1.pdf) | [AI1](downloads/3022/202504_3022_AI1.pdf) | [AI1](downloads/3022/202503_3022_AI1.pdf) | [AI1](downloads/3022/202502_3022_AI1.pdf) | [AI1](downloads/3022/202501_3022_AI1.pdf) | [AI1](downloads/3022/202404_3022_AI1.pdf) / [AI3](downloads/3022/202404_3022_AI3.pdf) | [AI1](downloads/3022/202403_3022_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3026 | 禾伸堂 | - | [AI1](downloads/3026/202601_3026_AI1.pdf) | [AI1](downloads/3026/202504_3026_AI1.pdf) | [AI1](downloads/3026/202503_3026_AI1.pdf) | [AI1](downloads/3026/202502_3026_AI1.pdf) | [AI1](downloads/3026/202501_3026_AI1.pdf) | [AI1](downloads/3026/202404_3026_AI1.pdf) / [AI3](downloads/3026/202404_3026_AI3.pdf) | [AI1](downloads/3026/202403_3026_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3029 | 零壹 | - | [AI1](downloads/3029/202601_3029_AI1.pdf) | [AI3](downloads/3029/202504_3029_AI3.pdf) | [AI1](downloads/3029/202503_3029_AI1.pdf) | [AI1](downloads/3029/202502_3029_AI1.pdf) | [AI1](downloads/3029/202501_3029_AI1.pdf) | [AI1](downloads/3029/202404_3029_AI1.pdf) / [AI3](downloads/3029/202404_3029_AI3.pdf) | [AI1](downloads/3029/202403_3029_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3034 | 聯詠 | - | [AI1](downloads/3034/202601_3034_AI1.pdf) | [AI1](downloads/3034/202504_3034_AI1.pdf) / [AI3](downloads/3034/202504_3034_AI3.pdf) | [AI1](downloads/3034/202503_3034_AI1.pdf) | [AI1](downloads/3034/202502_3034_AI1.pdf) | [AI1](downloads/3034/202501_3034_AI1.pdf) | [AI1](downloads/3034/202404_3034_AI1.pdf) / [AI3](downloads/3034/202404_3034_AI3.pdf) | [AI1](downloads/3034/202403_3034_AI1.pdf) | [AI1](downloads/3034/202402_3034_AI1.pdf) | [AI1](downloads/3034/202401_3034_AI1.pdf) | [AI1](downloads/3034/202304_3034_AI1.pdf) / [AI3](downloads/3034/202304_3034_AI3.pdf) | [AI1](downloads/3034/202303_3034_AI1.pdf) | [AI1](downloads/3034/202302_3034_AI1.pdf) | [AI1](downloads/3034/202301_3034_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3035 | 智原 | - | [AI1](downloads/3035/202601_3035_AI1.pdf) | [AI1](downloads/3035/202504_3035_AI1.pdf) | [AI1](downloads/3035/202503_3035_AI1.pdf) | [AI1](downloads/3035/202502_3035_AI1.pdf) | [AI1](downloads/3035/202501_3035_AI1.pdf) | [AI1](downloads/3035/202404_3035_AI1.pdf) / [AI3](downloads/3035/202404_3035_AI3.pdf) | [AI1](downloads/3035/202403_3035_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3045 | 台灣大 | - | [AI1](downloads/3045/202601_3045_AI1.pdf) | [AI1](downloads/3045/202504_3045_AI1.pdf) / [AI3](downloads/3045/202504_3045_AI3.pdf) | [AI1](downloads/3045/202503_3045_AI1.pdf) | [AI1](downloads/3045/202502_3045_AI1.pdf) | [AI1](downloads/3045/202501_3045_AI1.pdf) | [AI1](downloads/3045/202404_3045_AI1.pdf) / [AI3](downloads/3045/202404_3045_AI3.pdf) | [AI1](downloads/3045/202403_3045_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3048 | 益登 | - | [AI1](downloads/3048/202601_3048_AI1.pdf) | [AI1](downloads/3048/202504_3048_AI1.pdf) / [AI3](downloads/3048/202504_3048_AI3.pdf) | [AI1](downloads/3048/202503_3048_AI1.pdf) | [AI1](downloads/3048/202502_3048_AI1.pdf) | [AI1](downloads/3048/202501_3048_AI1.pdf) | [AI1](downloads/3048/202404_3048_AI1.pdf) / [AI3](downloads/3048/202404_3048_AI3.pdf) | [AI1](downloads/3048/202403_3048_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3150 | 鈺寶-創 | - | [AI2](downloads/3150/202601_3150_AI2.pdf) | [AI2](downloads/3150/202504_3150_AI2.pdf) | [AI2](downloads/3150/202503_3150_AI2.pdf) | [AI2](downloads/3150/202502_3150_AI2.pdf) | [AI2](downloads/3150/202501_3150_AI2.pdf) | [AI2](downloads/3150/202404_3150_AI2.pdf) | [AI2](downloads/3150/202403_3150_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3158 | 嘉實 | - | [AI1](downloads/3158/202601_3158_AI1.pdf) | [AI1](downloads/3158/202504_3158_AI1.pdf) / [AI3](downloads/3158/202504_3158_AI3.pdf) | [AI1](downloads/3158/202503_3158_AI1.pdf) | [AI1](downloads/3158/202502_3158_AI1.pdf) | [AI1](downloads/3158/202501_3158_AI1.pdf) | [AI1](downloads/3158/202404_3158_AI1.pdf) / [AI3](downloads/3158/202404_3158_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3231 | 緯創 | - | [AI1](downloads/3231/202601_3231_AI1.pdf) | [AI1](downloads/3231/202504_3231_AI1.pdf) / [AI3](downloads/3231/202504_3231_AI3.pdf) | [AI1](downloads/3231/202503_3231_AI1.pdf) | [AI1](downloads/3231/202502_3231_AI1.pdf) | [AI1](downloads/3231/202501_3231_AI1.pdf) | [AI1](downloads/3231/202404_3231_AI1.pdf) / [AI3](downloads/3231/202404_3231_AI3.pdf) | [AI1](downloads/3231/202403_3231_AI1.pdf) | [AI1](downloads/3231/202402_3231_AI1.pdf) | [AI1](downloads/3231/202401_3231_AI1.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3260 | 威剛 | - | [AI1](downloads/3260/202601_3260_AI1.pdf) | [AI1](downloads/3260/202504_3260_AI1.pdf) / [AI3](downloads/3260/202504_3260_AI3.pdf) | [AI1](downloads/3260/202503_3260_AI1.pdf) | [AI1](downloads/3260/202502_3260_AI1.pdf) | [AI1](downloads/3260/202501_3260_AI1.pdf) | [AI1](downloads/3260/202404_3260_AI1.pdf) / [AI3](downloads/3260/202404_3260_AI3.pdf) | [AI1](downloads/3260/202403_3260_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3293 | 鈊象 | - | [AI1](downloads/3293/202601_3293_AI1.pdf) | [AI1](downloads/3293/202504_3293_AI1.pdf) | [AI1](downloads/3293/202503_3293_AI1.pdf) | [AI1](downloads/3293/202502_3293_AI1.pdf) | [AI1](downloads/3293/202501_3293_AI1.pdf) | [AI1](downloads/3293/202404_3293_AI1.pdf) / [AI3](downloads/3293/202404_3293_AI3.pdf) | [AI1](downloads/3293/202403_3293_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3356 | 奇偶 | - | [AI1](downloads/3356/202601_3356_AI1.pdf) | [AI1](downloads/3356/202504_3356_AI1.pdf) | [AI1](downloads/3356/202503_3356_AI1.pdf) | [AI1](downloads/3356/202502_3356_AI1.pdf) | [AI1](downloads/3356/202501_3356_AI1.pdf) | [AI1](downloads/3356/202404_3356_AI1.pdf) / [AI3](downloads/3356/202404_3356_AI3.pdf) | [AI1](downloads/3356/202403_3356_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3467 | 台灣精材 | - | [AI2](downloads/3467/202601_3467_AI2.pdf) | [AI2](downloads/3467/202504_3467_AI2.pdf) | [AI2](downloads/3467/202503_3467_AI2.pdf) | [AI2](downloads/3467/202502_3467_AI2.pdf) | [AI2](downloads/3467/202501_3467_AI2.pdf) | [AI3](downloads/3467/202404_3467_AI3.pdf) | [AI2](downloads/3467/202403_3467_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 3558 | 神準 | - | [AI1](downloads/3558/202601_3558_AI1.pdf) | [AI1](downloads/3558/202504_3558_AI1.pdf) / [AI3](downloads/3558/202504_3558_AI3.pdf) | [AI1](downloads/3558/202503_3558_AI1.pdf) | [AI1](downloads/3558/202502_3558_AI1.pdf) | [AI1](downloads/3558/202501_3558_AI1.pdf) | - | [AI1](downloads/3558/202403_3558_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 4114 | 健喬 | - | [AI1](downloads/4114/202601_4114_AI1.pdf) | [AI1](downloads/4114/202504_4114_AI1.pdf) / [AI3](downloads/4114/202504_4114_AI3.pdf) | [AI1](downloads/4114/202503_4114_AI1.pdf) | [AI1](downloads/4114/202502_4114_AI1.pdf) | [AI1](downloads/4114/202501_4114_AI1.pdf) | - | [AI1](downloads/4114/202403_4114_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 4749 | 新應材 | - | [AI1](downloads/4749/202601_4749_AI1.pdf) | [AI1](downloads/4749/202504_4749_AI1.pdf) | [AI1](downloads/4749/202503_4749_AI1.pdf) | [AI1](downloads/4749/202502_4749_AI1.pdf) | [AI1](downloads/4749/202501_4749_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 4938 | 和碩 | - | [AI1](downloads/4938/202601_4938_AI1.pdf) | [AI1](downloads/4938/202504_4938_AI1.pdf) / [AI3](downloads/4938/202504_4938_AI3.pdf) | [AI1](downloads/4938/202503_4938_AI1.pdf) | [AI1](downloads/4938/202502_4938_AI1.pdf) | [AI1](downloads/4938/202501_4938_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 4953 | 緯軟 | - | [AI1](downloads/4953/202601_4953_AI1.pdf) | [AI1](downloads/4953/202504_4953_AI1.pdf) / [AI3](downloads/4953/202504_4953_AI3.pdf) | [AI1](downloads/4953/202503_4953_AI1.pdf) | [AI1](downloads/4953/202502_4953_AI1.pdf) | [AI1](downloads/4953/202501_4953_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5203 | 訊連 | - | [AI1](downloads/5203/202601_5203_AI1.pdf) | [AI1](downloads/5203/202504_5203_AI1.pdf) / [AI3](downloads/5203/202504_5203_AI3.pdf) | [AI1](downloads/5203/202503_5203_AI1.pdf) | [AI1](downloads/5203/202502_5203_AI1.pdf) | [AI1](downloads/5203/202501_5203_AI1.pdf) | - | [AI1](downloads/5203/202403_5203_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5269 | 祥碩 | - | [AI1](downloads/5269/202601_5269_AI1.pdf) | [AI1](downloads/5269/202504_5269_AI1.pdf) / [AI3](downloads/5269/202504_5269_AI3.pdf) | [AI1](downloads/5269/202503_5269_AI1.pdf) | [AI1](downloads/5269/202502_5269_AI1.pdf) | [AI2](downloads/5269/202501_5269_AI2.pdf) | - | [AI2](downloads/5269/202403_5269_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5274 | 信驊 | - | [AI1](downloads/5274/202601_5274_AI1.pdf) | [AI1](downloads/5274/202504_5274_AI1.pdf) | [AI1](downloads/5274/202503_5274_AI1.pdf) | [AI1](downloads/5274/202502_5274_AI1.pdf) | [AI1](downloads/5274/202501_5274_AI1.pdf) | - | [AI1](downloads/5274/202403_5274_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5434 | 崇越 | - | [AI1](downloads/5434/202601_5434_AI1.pdf) | [AI1](downloads/5434/202504_5434_AI1.pdf) / [AI3](downloads/5434/202504_5434_AI3.pdf) | [AI1](downloads/5434/202503_5434_AI1.pdf) | [AI1](downloads/5434/202502_5434_AI1.pdf) | [AI1](downloads/5434/202501_5434_AI1.pdf) | - | [AI1](downloads/5434/202403_5434_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5536 | 聖暉 | - | [AI1](downloads/5536/202601_5536_AI1.pdf) | [AI1](downloads/5536/202504_5536_AI1.pdf) / [AI3](downloads/5536/202504_5536_AI3.pdf) | [AI1](downloads/5536/202503_5536_AI1.pdf) | [AI1](downloads/5536/202502_5536_AI1.pdf) | [AI1](downloads/5536/202501_5536_AI1.pdf) | - | [AI1](downloads/5536/202403_5536_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 5904 | 寶雅 | - | [AI2](downloads/5904/202601_5904_AI2.pdf) | [AI2](downloads/5904/202504_5904_AI2.pdf) | [AI2](downloads/5904/202503_5904_AI2.pdf) | [AI2](downloads/5904/202502_5904_AI2.pdf) | [AI2](downloads/5904/202501_5904_AI2.pdf) | - | [AI2](downloads/5904/202403_5904_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6035 | 悠遊卡 | - | - | - | - | [AI2](downloads/6035/202502_6035_AI2.pdf) | - | [AI2](downloads/6035/202404_6035_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6123 | 上奇 | - | [AI1](downloads/6123/202601_6123_AI1.pdf) | [AI1](downloads/6123/202504_6123_AI1.pdf) | [AI1](downloads/6123/202503_6123_AI1.pdf) | [AI1](downloads/6123/202502_6123_AI1.pdf) | [AI1](downloads/6123/202501_6123_AI1.pdf) | - | [AI1](downloads/6123/202403_6123_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6125 | 廣運 | - | [AI1](downloads/6125/202601_6125_AI1.pdf) | [AI1](downloads/6125/202504_6125_AI1.pdf) | [AI1](downloads/6125/202503_6125_AI1.pdf) | [AI1](downloads/6125/202502_6125_AI1.pdf) | [AI1](downloads/6125/202501_6125_AI1.pdf) | [AI1](downloads/6125/202404_6125_AI1.pdf) / [AI3](downloads/6125/202404_6125_AI3.pdf) | [AI1](downloads/6125/202403_6125_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6182 | 合晶 | - | [AI1](downloads/6182/202601_6182_AI1.pdf) | [AI1](downloads/6182/202504_6182_AI1.pdf) / [AI3](downloads/6182/202504_6182_AI3.pdf) | [AI1](downloads/6182/202503_6182_AI1.pdf) | [AI1](downloads/6182/202502_6182_AI1.pdf) | [AI1](downloads/6182/202501_6182_AI1.pdf) | [AI1](downloads/6182/202404_6182_AI1.pdf) / [AI3](downloads/6182/202404_6182_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6214 | 精誠 | - | [AI1](downloads/6214/202601_6214_AI1.pdf) | [AI1](downloads/6214/202504_6214_AI1.pdf) / [AI3](downloads/6214/202504_6214_AI3.pdf) | [AI1](downloads/6214/202503_6214_AI1.pdf) | [AI1](downloads/6214/202502_6214_AI1.pdf) | [AI1](downloads/6214/202501_6214_AI1.pdf) | [AI1](downloads/6214/202404_6214_AI1.pdf) / [AI3](downloads/6214/202404_6214_AI3.pdf) | [AI1](downloads/6214/202403_6214_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6231 | 系微 | - | [AI1](downloads/6231/202601_6231_AI1.pdf) | [AI3](downloads/6231/202504_6231_AI3.pdf) | [AI1](downloads/6231/202503_6231_AI1.pdf) | [AI1](downloads/6231/202502_6231_AI1.pdf) | [AI1](downloads/6231/202501_6231_AI1.pdf) | [AI1](downloads/6231/202404_6231_AI1.pdf) / [AI3](downloads/6231/202404_6231_AI3.pdf) | [AI1](downloads/6231/202403_6231_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6285 | 啟碁 | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6425 | 易發 | - | [AI1](downloads/6425/202601_6425_AI1.pdf) | [AI1](downloads/6425/202504_6425_AI1.pdf) / [AI3](downloads/6425/202504_6425_AI3.pdf) | [AI1](downloads/6425/202503_6425_AI1.pdf) | [AI1](downloads/6425/202502_6425_AI1.pdf) | [AI1](downloads/6425/202501_6425_AI1.pdf) | [AI1](downloads/6425/202404_6425_AI1.pdf) / [AI3](downloads/6425/202404_6425_AI3.pdf) | [AI1](downloads/6425/202403_6425_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6442 | 光聖 | - | [AI1](downloads/6442/202601_6442_AI1.pdf) | [AI1](downloads/6442/202504_6442_AI1.pdf) / [AI3](downloads/6442/202504_6442_AI3.pdf) | [AI1](downloads/6442/202503_6442_AI1.pdf) | [AI1](downloads/6442/202502_6442_AI1.pdf) | [AI1](downloads/6442/202501_6442_AI1.pdf) | [AI1](downloads/6442/202404_6442_AI1.pdf) / [AI3](downloads/6442/202404_6442_AI3.pdf) | [AI1](downloads/6442/202403_6442_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6462 | 神盾 | - | [AI1](downloads/6462/202601_6462_AI1.pdf) | [AI1](downloads/6462/202504_6462_AI1.pdf) / [AI3](downloads/6462/202504_6462_AI3.pdf) | [AI1](downloads/6462/202503_6462_AI1.pdf) | [AI1](downloads/6462/202502_6462_AI1.pdf) | [AI1](downloads/6462/202501_6462_AI1.pdf) | [AI1](downloads/6462/202404_6462_AI1.pdf) / [AI3](downloads/6462/202404_6462_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6506 | 雙邦 | - | [AI1](downloads/6506/202601_6506_AI1.pdf) | [AI1](downloads/6506/202504_6506_AI1.pdf) | [AI1](downloads/6506/202503_6506_AI1.pdf) | [AI1](downloads/6506/202502_6506_AI1.pdf) | [AI1](downloads/6506/202501_6506_AI1.pdf) | [AI1](downloads/6506/202404_6506_AI1.pdf) / [AI3](downloads/6506/202404_6506_AI3.pdf) | [AI1](downloads/6506/202403_6506_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6510 | 精測 | - | [AI1](downloads/6510/202601_6510_AI1.pdf) | [AI1](downloads/6510/202504_6510_AI1.pdf) / [AI3](downloads/6510/202504_6510_AI3.pdf) | [AI1](downloads/6510/202503_6510_AI1.pdf) | [AI1](downloads/6510/202502_6510_AI1.pdf) | [AI1](downloads/6510/202501_6510_AI1.pdf) | [AI1](downloads/6510/202404_6510_AI1.pdf) / [AI3](downloads/6510/202404_6510_AI3.pdf) | [AI1](downloads/6510/202403_6510_AI1.pdf) | [AI1](downloads/6510/202402_6510_AI1.pdf) | [AI1](downloads/6510/202401_6510_AI1.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6526 | 達發 | - | [AI1](downloads/6526/202601_6526_AI1.pdf) | [AI1](downloads/6526/202504_6526_AI1.pdf) / [AI3](downloads/6526/202504_6526_AI3.pdf) | [AI1](downloads/6526/202503_6526_AI1.pdf) | [AI1](downloads/6526/202502_6526_AI1.pdf) | [AI1](downloads/6526/202501_6526_AI1.pdf) | [AI1](downloads/6526/202404_6526_AI1.pdf) / [AI3](downloads/6526/202404_6526_AI3.pdf) | [AI1](downloads/6526/202403_6526_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6561 | 是方 | - | [AI1](downloads/6561/202601_6561_AI1.pdf) | [AI1](downloads/6561/202504_6561_AI1.pdf) / [AI3](downloads/6561/202504_6561_AI3.pdf) | [AI1](downloads/6561/202503_6561_AI1.pdf) | [AI1](downloads/6561/202502_6561_AI1.pdf) | [AI1](downloads/6561/202501_6561_AI1.pdf) | [AI1](downloads/6561/202404_6561_AI1.pdf) / [AI3](downloads/6561/202404_6561_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6597 | 立誠 | - | [AI2](downloads/6597/202601_6597_AI2.pdf) | [AI2](downloads/6597/202504_6597_AI2.pdf) | [AI2](downloads/6597/202503_6597_AI2.pdf) | [AI2](downloads/6597/202502_6597_AI2.pdf) | [AI2](downloads/6597/202501_6597_AI2.pdf) | [AI2](downloads/6597/202404_6597_AI2.pdf) | [AI2](downloads/6597/202403_6597_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6613 | 朋億 | - | [AI1](downloads/6613/202601_6613_AI1.pdf) | [AI1](downloads/6613/202504_6613_AI1.pdf) | [AI1](downloads/6613/202503_6613_AI1.pdf) | [AI1](downloads/6613/202502_6613_AI1.pdf) | [AI1](downloads/6613/202501_6613_AI1.pdf) | [AI1](downloads/6613/202404_6613_AI1.pdf) / [AI3](downloads/6613/202404_6613_AI3.pdf) | [AI1](downloads/6613/202403_6613_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6669 | 緯穎 | - | [AI1](downloads/6669/202601_6669_AI1.pdf) | [AI1](downloads/6669/202504_6669_AI1.pdf) / [AI3](downloads/6669/202504_6669_AI3.pdf) | [AI1](downloads/6669/202503_6669_AI1.pdf) | [AI1](downloads/6669/202502_6669_AI1.pdf) | [AI1](downloads/6669/202501_6669_AI1.pdf) | [AI1](downloads/6669/202404_6669_AI1.pdf) / [AI3](downloads/6669/202404_6669_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6690 | 安碁資訊 | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6695 | 芯鼎 | - | [AI1](downloads/6695/202601_6695_AI1.pdf) | [AI1](downloads/6695/202504_6695_AI1.pdf) | [AI1](downloads/6695/202503_6695_AI1.pdf) | [AI1](downloads/6695/202502_6695_AI1.pdf) | [AI1](downloads/6695/202501_6695_AI1.pdf) | [AI1](downloads/6695/202404_6695_AI1.pdf) / [AI3](downloads/6695/202404_6695_AI3.pdf) | [AI1](downloads/6695/202403_6695_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6699 | 奇邑 | - | - | - | - | [AI1](downloads/6699/202502_6699_AI1.pdf) | - | [AI1](downloads/6699/202404_6699_AI1.pdf) / [AI3](downloads/6699/202404_6699_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6720 | 久昌 | - | [AI1](downloads/6720/202601_6720_AI1.pdf) | [AI1](downloads/6720/202504_6720_AI1.pdf) / [AI3](downloads/6720/202504_6720_AI3.pdf) | [AI1](downloads/6720/202503_6720_AI1.pdf) | [AI1](downloads/6720/202502_6720_AI1.pdf) | [AI1](downloads/6720/202501_6720_AI1.pdf) | [AI1](downloads/6720/202404_6720_AI1.pdf) / [AI3](downloads/6720/202404_6720_AI3.pdf) | [AI1](downloads/6720/202403_6720_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6751 | 智聯服務 | - | [AI1](downloads/6751/202601_6751_AI1.pdf) | [AI1](downloads/6751/202504_6751_AI1.pdf) / [AI3](downloads/6751/202504_6751_AI3.pdf) | [AI1](downloads/6751/202503_6751_AI1.pdf) | [AI1](downloads/6751/202502_6751_AI1.pdf) | [AI1](downloads/6751/202501_6751_AI1.pdf) | [AI1](downloads/6751/202404_6751_AI1.pdf) | [AI1](downloads/6751/202403_6751_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6757 | 台灣虎航 | - | [AI2](downloads/6757/202601_6757_AI2.pdf) | [AI2](downloads/6757/202504_6757_AI2.pdf) | [AI2](downloads/6757/202503_6757_AI2.pdf) | [AI2](downloads/6757/202502_6757_AI2.pdf) | [AI2](downloads/6757/202501_6757_AI2.pdf) | [AI2](downloads/6757/202404_6757_AI2.pdf) | [AI2](downloads/6757/202403_6757_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6763 | 綠界科技 | - | [AI1](downloads/6763/202601_6763_AI1.pdf) | [AI3](downloads/6763/202504_6763_AI3.pdf) | [AI1](downloads/6763/202503_6763_AI1.pdf) | [AI1](downloads/6763/202502_6763_AI1.pdf) | [AI1](downloads/6763/202501_6763_AI1.pdf) | [AI1](downloads/6763/202404_6763_AI1.pdf) / [AI3](downloads/6763/202404_6763_AI3.pdf) | [AI1](downloads/6763/202403_6763_AI1.pdf) | [AI1](downloads/6763/202402_6763_AI1.pdf) | [AI1](downloads/6763/202401_6763_AI1.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6811 | 宏碁資訊 | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6850 | 光鼎生技 | - | - | - | - | [AI1](downloads/6850/202502_6850_AI1.pdf) | - | [AI1](downloads/6850/202404_6850_AI1.pdf) / [AI3](downloads/6850/202404_6850_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6902 | GOGOLOOK | - | [AI1](downloads/6902/202601_6902_AI1.pdf) | [AI1](downloads/6902/202504_6902_AI1.pdf) / [AI3](downloads/6902/202504_6902_AI3.pdf) | [AI1](downloads/6902/202503_6902_AI1.pdf) | [AI1](downloads/6902/202502_6902_AI1.pdf) | [AI1](downloads/6902/202501_6902_AI1.pdf) | [AI1](downloads/6902/202404_6902_AI1.pdf) / [AI3](downloads/6902/202404_6902_AI3.pdf) | [AI1](downloads/6902/202403_6902_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6918 | 愛派司 | - | [AI1](downloads/6918/202601_6918_AI1.pdf) | [AI1](downloads/6918/202504_6918_AI1.pdf) / [AI3](downloads/6918/202504_6918_AI3.pdf) | [AI1](downloads/6918/202503_6918_AI1.pdf) | [AI1](downloads/6918/202502_6918_AI1.pdf) | [AI1](downloads/6918/202501_6918_AI1.pdf) | - | [AI1](downloads/6918/202403_6918_AI1.pdf) / [AI4](downloads/6918/202403_6918_AI4.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6925 | 意藍 | - | [AI2](downloads/6925/202601_6925_AI2.pdf) | [AI2](downloads/6925/202504_6925_AI2.pdf) | [AI2](downloads/6925/202503_6925_AI2.pdf) | [AI2](downloads/6925/202502_6925_AI2.pdf) | [AI2](downloads/6925/202501_6925_AI2.pdf) | - | [AI2](downloads/6925/202403_6925_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6962 | 奕力-KY | - | [AI1](downloads/6962/202601_6962_AI1.pdf) | [AI1](downloads/6962/202504_6962_AI1.pdf) | [AI1](downloads/6962/202503_6962_AI1.pdf) | - | [AI1](downloads/6962/202501_6962_AI1.pdf) | - | [AI1](downloads/6962/202403_6962_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6996 | 力領科技 | - | [AI2](downloads/6996/202601_6996_AI2.pdf) | [AI2](downloads/6996/202504_6996_AI2.pdf) | [AI2](downloads/6996/202503_6996_AI2.pdf) | [AI2](downloads/6996/202502_6996_AI2.pdf) | [AI2](downloads/6996/202501_6996_AI2.pdf) | [AI2](downloads/6996/202404_6996_AI2.pdf) | [AI2](downloads/6996/202403_6996_AI2.pdf) | [AI2](downloads/6996/202402_6996_AI2.pdf) | [AI2](downloads/6996/202401_6996_AI2.pdf) | [AI2](downloads/6996/202304_6996_AI2.pdf) | - | [AI2](downloads/6996/202302_6996_AI2.pdf) | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 6997 | 博弘 | - | [AI1](downloads/6997/202601_6997_AI1.pdf) | [AI1](downloads/6997/202504_6997_AI1.pdf) / [AI3](downloads/6997/202504_6997_AI3.pdf) | [AI1](downloads/6997/202503_6997_AI1.pdf) | [AI1](downloads/6997/202502_6997_AI1.pdf) | [AI1](downloads/6997/202501_6997_AI1.pdf) | - | [AI1](downloads/6997/202403_6997_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7547 | 碩網 | - | [AI1](downloads/7547/202601_7547_AI1.pdf) | [AI1](downloads/7547/202504_7547_AI1.pdf) / [AI3](downloads/7547/202504_7547_AI3.pdf) | [AI1](downloads/7547/202503_7547_AI1.pdf) | [AI1](downloads/7547/202502_7547_AI1.pdf) | [AI1](downloads/7547/202501_7547_AI1.pdf) | - | [AI1](downloads/7547/202403_7547_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7703 | 銳澤 | - | [AI1](downloads/7703/202601_7703_AI1.pdf) | [AI1](downloads/7703/202504_7703_AI1.pdf) / [AI3](downloads/7703/202504_7703_AI3.pdf) | [AI1](downloads/7703/202503_7703_AI1.pdf) | [AI1](downloads/7703/202502_7703_AI1.pdf) | [AI1](downloads/7703/202501_7703_AI1.pdf) | - | [AI1](downloads/7703/202403_7703_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7704 | 明遠精密 | - | [AI1](downloads/7704/202601_7704_AI1.pdf) | [AI1](downloads/7704/202504_7704_AI1.pdf) | [AI1](downloads/7704/202503_7704_AI1.pdf) | [AI1](downloads/7704/202502_7704_AI1.pdf) | [AI1](downloads/7704/202501_7704_AI1.pdf) | - | [AI1](downloads/7704/202403_7704_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7705 | 三商餐飲 | - | [AI1](downloads/7705/202601_7705_AI1.pdf) | [AI1](downloads/7705/202504_7705_AI1.pdf) | [AI1](downloads/7705/202503_7705_AI1.pdf) | [AI1](downloads/7705/202502_7705_AI1.pdf) | [AI1](downloads/7705/202501_7705_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7708 | 全家餐飲 | - | [AI1](downloads/7708/202601_7708_AI1.pdf) | [AI1](downloads/7708/202504_7708_AI1.pdf) / [AI3](downloads/7708/202504_7708_AI3.pdf) | [AI1](downloads/7708/202503_7708_AI1.pdf) | [AI1](downloads/7708/202502_7708_AI1.pdf) | [AI1](downloads/7708/202501_7708_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7709 | 榮田 | - | [AI2](downloads/7709/202601_7709_AI2.pdf) | [AI2](downloads/7709/202504_7709_AI2.pdf) | [AI2](downloads/7709/202503_7709_AI2.pdf) | [AI2](downloads/7709/202502_7709_AI2.pdf) | [AI2](downloads/7709/202501_7709_AI2.pdf) | [AI2](downloads/7709/202404_7709_AI2.pdf) | [AI2](downloads/7709/202403_7709_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7712 | 博盛半導體 | - | [AI1](downloads/7712/202601_7712_AI1.pdf) | [AI1](downloads/7712/202504_7712_AI1.pdf) / [AI3](downloads/7712/202504_7712_AI3.pdf) | [AI1](downloads/7712/202503_7712_AI1.pdf) | [AI1](downloads/7712/202502_7712_AI1.pdf) | [AI1](downloads/7712/202501_7712_AI1.pdf) | [AI1](downloads/7712/202404_7712_AI1.pdf) / [AI3](downloads/7712/202404_7712_AI3.pdf) | [AI1](downloads/7712/202403_7712_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7713 | 威力德生醫 | - | [AI1](downloads/7713/202601_7713_AI1.pdf) | [AI1](downloads/7713/202504_7713_AI1.pdf) / [AI3](downloads/7713/202504_7713_AI3.pdf) | [AI1](downloads/7713/202503_7713_AI1.pdf) | [AI1](downloads/7713/202502_7713_AI1.pdf) | [AI1](downloads/7713/202501_7713_AI1.pdf) | [AI1](downloads/7713/202404_7713_AI1.pdf) / [AI3](downloads/7713/202404_7713_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7722 | LINEPAY | - | [AI1](downloads/7722/202601_7722_AI1.pdf) | [AI1](downloads/7722/202504_7722_AI1.pdf) | [AI1](downloads/7722/202503_7722_AI1.pdf) | [AI1](downloads/7722/202502_7722_AI1.pdf) | [AI1](downloads/7722/202501_7722_AI1.pdf) | [AI1](downloads/7722/202404_7722_AI1.pdf) / [AI3](downloads/7722/202404_7722_AI3.pdf) | [AI1](downloads/7722/202403_7722_AI1.pdf) | [AI1](downloads/7722/202402_7722_AI1.pdf) | [AI1](downloads/7722/202401_7722_AI1.pdf) | [AI1](downloads/7722/202304_7722_AI1.pdf) / [AI3](downloads/7722/202304_7722_AI3.pdf) | - | [AI2](downloads/7722/202302_7722_AI2.pdf) | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7728 | 光焱科技 | - | [AI1](downloads/7728/202601_7728_AI1.pdf) | [AI1](downloads/7728/202504_7728_AI1.pdf) / [AI3](downloads/7728/202504_7728_AI3.pdf) | [AI1](downloads/7728/202503_7728_AI1.pdf) | [AI1](downloads/7728/202502_7728_AI1.pdf) | [AI1](downloads/7728/202501_7728_AI1.pdf) | [AI1](downloads/7728/202404_7728_AI1.pdf) / [AI3](downloads/7728/202404_7728_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7732 | 金興精密 | - | [AI1](downloads/7732/202601_7732_AI1.pdf) | [AI1](downloads/7732/202504_7732_AI1.pdf) / [AI3](downloads/7732/202504_7732_AI3.pdf) | [AI1](downloads/7732/202503_7732_AI1.pdf) | [AI1](downloads/7732/202502_7732_AI1.pdf) | [AI1](downloads/7732/202501_7732_AI1.pdf) | [AI1](downloads/7732/202404_7732_AI1.pdf) / [AI3](downloads/7732/202404_7732_AI3.pdf) | [AI1](downloads/7732/202403_7732_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7734 | 印能科技 | - | [AI1](downloads/7734/202601_7734_AI1.pdf) | [AI1](downloads/7734/202504_7734_AI1.pdf) / [AI3](downloads/7734/202504_7734_AI3.pdf) | [AI1](downloads/7734/202503_7734_AI1.pdf) | [AI1](downloads/7734/202502_7734_AI1.pdf) | [AI1](downloads/7734/202501_7734_AI1.pdf) | [AI1](downloads/7734/202404_7734_AI1.pdf) / [AI3](downloads/7734/202404_7734_AI3.pdf) | [AI1](downloads/7734/202403_7734_AI1.pdf) | [AI1](downloads/7734/202402_7734_AI1.pdf) | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7736 | 虎山 | - | [AI1](downloads/7736/202601_7736_AI1.pdf) | [AI1](downloads/7736/202504_7736_AI1.pdf) / [AI3](downloads/7736/202504_7736_AI3.pdf) | [AI1](downloads/7736/202503_7736_AI1.pdf) | [AI1](downloads/7736/202502_7736_AI1.pdf) | [AI1](downloads/7736/202501_7736_AI1.pdf) | [AI1](downloads/7736/202404_7736_AI1.pdf) / [AI3](downloads/7736/202404_7736_AI3.pdf) | [AI1](downloads/7736/202403_7736_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7737 | 凱鈿 | - | - | - | - | [AI1](downloads/7737/202502_7737_AI1.pdf) | - | [AI1](downloads/7737/202404_7737_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7747 | 昕奇雲端 | - | [AI1](downloads/7747/202601_7747_AI1.pdf) | [AI1](downloads/7747/202504_7747_AI1.pdf) | [AI1](downloads/7747/202503_7747_AI1.pdf) | [AI1](downloads/7747/202502_7747_AI1.pdf) | [AI1](downloads/7747/202501_7747_AI1.pdf) | [AI1](downloads/7747/202404_7747_AI1.pdf) | [AI1](downloads/7747/202403_7747_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7749 | 意騰-KY | - | [AI1](downloads/7749/202601_7749_AI1.pdf) | [AI1](downloads/7749/202504_7749_AI1.pdf) | [AI1](downloads/7749/202503_7749_AI1.pdf) | - | [AI1](downloads/7749/202501_7749_AI1.pdf) | [AI1](downloads/7749/202404_7749_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7765 | 中華資安 | - | [AI2](downloads/7765/202601_7765_AI2.pdf) | [AI2](downloads/7765/202504_7765_AI2.pdf) | [AI2](downloads/7765/202503_7765_AI2.pdf) | [AI2](downloads/7765/202502_7765_AI2.pdf) | [AI2](downloads/7765/202501_7765_AI2.pdf) | [AI2](downloads/7765/202404_7765_AI2.pdf) | - | [AI2](downloads/7765/202402_7765_AI2.pdf) | - | [AI2](downloads/7765/202304_7765_AI2.pdf) | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7769 | 鴻勁 | - | [AI1](downloads/7769/202601_7769_AI1.pdf) | [AI1](downloads/7769/202504_7769_AI1.pdf) / [AI3](downloads/7769/202504_7769_AI3.pdf) | [AI1](downloads/7769/202503_7769_AI1.pdf) | [AI1](downloads/7769/202502_7769_AI1.pdf) | [AI1](downloads/7769/202501_7769_AI1.pdf) | [AI1](downloads/7769/202404_7769_AI1.pdf) / [AI2](downloads/7769/202404_7769_AI2.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7794 | 宏碁智新 | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 7805 | 威聯通 | - | [AI1](downloads/7805/202601_7805_AI1.pdf) | [AI1](downloads/7805/202504_7805_AI1.pdf) / [AI3](downloads/7805/202504_7805_AI3.pdf) | [AI1](downloads/7805/202503_7805_AI1.pdf) | [AI1](downloads/7805/202502_7805_AI1.pdf) | [AI1](downloads/7805/202501_7805_AI1.pdf) | [AI1](downloads/7805/202404_7805_AI1.pdf) / [AI3](downloads/7805/202404_7805_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 8016 | 矽創 | - | [AI1](downloads/8016/202601_8016_AI1.pdf) | [AI1](downloads/8016/202504_8016_AI1.pdf) / [AI3](downloads/8016/202504_8016_AI3.pdf) | [AI1](downloads/8016/202503_8016_AI1.pdf) | [AI1](downloads/8016/202502_8016_AI1.pdf) | [AI1](downloads/8016/202501_8016_AI1.pdf) | [AI1](downloads/8016/202404_8016_AI1.pdf) / [AI3](downloads/8016/202404_8016_AI3.pdf) | [AI1](downloads/8016/202403_8016_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 8045 | 達運光電 | - | [AI1](downloads/8045/202601_8045_AI1.pdf) | [AI1](downloads/8045/202504_8045_AI1.pdf) | [AI1](downloads/8045/202503_8045_AI1.pdf) | [AI1](downloads/8045/202502_8045_AI1.pdf) | [AI1](downloads/8045/202501_8045_AI1.pdf) | [AI1](downloads/8045/202404_8045_AI1.pdf) / [AI3](downloads/8045/202404_8045_AI3.pdf) | [AI1](downloads/8045/202403_8045_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 8272 | 全景軟體 | - | [AI2](downloads/8272/202601_8272_AI2.pdf) | [AI2](downloads/8272/202504_8272_AI2.pdf) | [AI2](downloads/8272/202503_8272_AI2.pdf) | [AI2](downloads/8272/202502_8272_AI2.pdf) | [AI2](downloads/8272/202501_8272_AI2.pdf) | [AI2](downloads/8272/202404_8272_AI2.pdf) | [AI2](downloads/8272/202403_8272_AI2.pdf) | [AI2](downloads/8272/202402_8272_AI2.pdf) | [AI2](downloads/8272/202401_8272_AI2.pdf) | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 8299 | 群聯 | - | [AI1](downloads/8299/202601_8299_AI1.pdf) | [AI1](downloads/8299/202504_8299_AI1.pdf) | [AI1](downloads/8299/202503_8299_AI1.pdf) | [AI1](downloads/8299/202502_8299_AI1.pdf) | [AI1](downloads/8299/202501_8299_AI1.pdf) | [AI1](downloads/8299/202404_8299_AI1.pdf) / [AI3](downloads/8299/202404_8299_AI3.pdf) | [AI1](downloads/8299/202403_8299_AI1.pdf) | [AI1](downloads/8299/202402_8299_AI1.pdf) | [AI1](downloads/8299/202401_8299_AI1.pdf) | [AI1](downloads/8299/202304_8299_AI1.pdf) / [AI3](downloads/8299/202304_8299_AI3.pdf) | [AI1](downloads/8299/202303_8299_AI1.pdf) | [AI1](downloads/8299/202302_8299_AI1.pdf) | [AI1](downloads/8299/202301_8299_AI1.pdf) | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 8454 | 富邦媒 | - | [AI1](downloads/8454/202601_8454_AI1.pdf) | [AI1](downloads/8454/202504_8454_AI1.pdf) / [AI3](downloads/8454/202504_8454_AI3.pdf) | [AI1](downloads/8454/202503_8454_AI1.pdf) | [AI1](downloads/8454/202502_8454_AI1.pdf) | [AI1](downloads/8454/202501_8454_AI1.pdf) | [AI1](downloads/8454/202404_8454_AI1.pdf) / [AI3](downloads/8454/202404_8454_AI3.pdf) | [AI1](downloads/8454/202403_8454_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 9914 | 美利達 | - | [AI1](downloads/9914/202601_9914_AI1.pdf) | [AI1](downloads/9914/202504_9914_AI1.pdf) / [AI3](downloads/9914/202504_9914_AI3.pdf) | [AI1](downloads/9914/202503_9914_AI1.pdf) | [AI1](downloads/9914/202502_9914_AI1.pdf) | [AI1](downloads/9914/202501_9914_AI1.pdf) | [AI1](downloads/9914/202404_9914_AI1.pdf) / [AI3](downloads/9914/202404_9914_AI3.pdf) | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 9917 | 中保科 | - | [AI1](downloads/9917/202601_9917_AI1.pdf) | [AI1](downloads/9917/202504_9917_AI1.pdf) / [AI3](downloads/9917/202504_9917_AI3.pdf) | [AI1](downloads/9917/202503_9917_AI1.pdf) | [AI1](downloads/9917/202502_9917_AI1.pdf) | [AI1](downloads/9917/202501_9917_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |
| 9921 | 巨大 | - | [AI1](downloads/9921/202601_9921_AI1.pdf) | [AI1](downloads/9921/202504_9921_AI1.pdf) / [AI3](downloads/9921/202504_9921_AI3.pdf) | [AI1](downloads/9921/202503_9921_AI1.pdf) | [AI1](downloads/9921/202502_9921_AI1.pdf) | [AI1](downloads/9921/202501_9921_AI1.pdf) | - | - | - | - | - | - | - | - | - | - | - | - | 2026-06-09T13:58:06.341789+08:00 |

---

## 📊 Current Download Status

> **Last Updated**: 2026-07-02 | **Source**: `mops_matrix_latest.csv`

**118 companies tracked**

### 季財報 概況

| Quarter | 季財報 | Coverage | Notes |
|---------|--------|----------|-------|
| 2026 Q2 | 0 / 118 | — |  |
| 2026 Q1 | 109 / 118 | 92% | Filing deadline: May 15 |
| 2025 Q4 | 109 / 118 | 92% | Filing deadline: Mar 31 (next year) |
| 2025 Q3 | 109 / 118 | 92% |  |
| 2025 Q2 | 111 / 118 | 94% |  |
| 2025 Q1 | 105 / 118 | 89% |  |
| 2024 Q4 | 90 / 118 | 76% |  |
| 2024 Q3 | 78 / 118 | 66% |  |
| 2024 Q2 | 17 / 118 | 14% |  |
| 2024 Q1 | 15 / 118 | 13% |  |
| 2023 Q4 | 12 / 118 | 10% |  |
| 2023 Q3 | 9 / 118 | 8% |  |
| 2023 Q2 | 11 / 118 | 9% |  |
| 2023 Q1 | 9 / 118 | 8% |  |
| 2020 Q4 | 1 / 118 | 1% |  |
| 2020 Q3 | 1 / 118 | 1% |  |
| 2020 Q2 | 1 / 118 | 1% |  |
| 2020 Q1 | 1 / 118 | 1% |  |
| process_timestamp | 118 / 118 | 100% |  |

---

### 📂 法說會 & 新聞 資料庫

#### 完成度概況

> 各公司法說會簡報、逐字稿、新聞收錄數量。點擊公司名稱展開詳細連結。

| 公司 | 法說會 PDF/MD | 逐字稿 | 新聞 |
|------|:------------:|:------:|:----:|
| [2357 華碩](#2357-華碩) | 11 季 | 2 | 2 |
| [2382 廣達](#2382-廣達) | — | 1 | — |

**覆蓋率**：2 / 118 companies

---

#### 2357 華碩

<details>
<summary>法說會（季度）</summary>

| Quarter | 法說會 PDF/MD | 逐字稿 |
|---------|:------------:|:------:|
| 2025 Q3 | [MD](downloads/2357/InvestorRelation/2025Q3_IR_Chinese.md) | [2025-11-11](downloads/2357/InvestorRelation/法說會逐字稿/華碩_2025-11-11.md) / [2025-12-08](downloads/2357/InvestorRelation/法說會逐字稿/華碩_2025-12-08.md) |
| 2025 Q2 | [MD](downloads/2357/InvestorRelation/2025Q2_IR_Chinese.md) | — |
| 2025 Q1 | [MD](downloads/2357/InvestorRelation/2025Q1_IR_Chinese.md) | — |
| 2024 Q4 | [MD](downloads/2357/InvestorRelation/2024Q4_IR_Chinese.md) | — |
| 2024 Q3 | [MD](downloads/2357/InvestorRelation/2024Q3_IR_Chinese.md) | — |
| 2024 Q2 | [MD](downloads/2357/InvestorRelation/2024Q2_IR_Chinese.md) | — |
| 2024 Q1 | [MD](downloads/2357/InvestorRelation/2024Q1_IR_Chinese.md) | — |
| 2023 Q4 | [MD](downloads/2357/InvestorRelation/2023Q4_IR_Chinese.md) | — |
| 2023 Q3 | [MD](downloads/2357/InvestorRelation/2023Q3_IR_Chinese.md) | — |
| 2023 Q2 | [MD](downloads/2357/InvestorRelation/2023Q2_IR_Chinese.md) | — |
| 2023 Q1 | [MD](downloads/2357/InvestorRelation/2023Q1_IR_Chinese.md) | — |

</details>

<details>
<summary>新聞</summary>

| 日期 | 標題 |
|------|------|
| 2026-01-29 | [華碩全力衝 AI 伺服器 2026年獨立事業群「福將」朱培蘭掌旗](downloads/2357/News/news_20260129_udn_asus_server_bg.md) |
| 2026-03-06 | [華碩 AI 伺服器戰略不攻大廠 北美四大 CSP 之外生意空間有多大？](downloads/2357/News/2026-03-06_udn_asus-ai-server-tier2-csp.md) |

</details>

---

#### 2382 廣達

<!-- END_STATUS -->

### Report Types (2025 Q3)

| Type | Count |
|------|-------|
| AI1 | 98 |
| AI2 | 11 |

### Companies Missing Recent Data

**Missing 2025 Q3** (9 companies):
2353 宏碁、6035 悠遊卡、6285 啟碁、6690 安碁資訊、6699 奇邑、6811 宏碁資訊、6850 光鼎生技、7737 凱鈿、7794 宏碁智新

**Missing 2025 Q2** (7 companies):
2353 宏碁、6285 啟碁、6690 安碁資訊、6811 宏碁資訊、6962 奕力-KY、7749 意騰-KY、7794 宏碁智新

**Missing 2025 Q1** (13 companies):
2345 智邦、2353 宏碁、2359 所羅門、2383 台光電、2405 輔信、6035 悠遊卡、6285 啟碁、6690 安碁資訊、6699 奇邑、6811 宏碁資訊、6850 光鼎生技、7737 凱鈿、7794 宏碁智新

---

## 🎯 What This Tool Does

- **Automates MOPS Downloads**: Fetches IFRSs financial reports in Chinese format from Taiwan's official MOPS system
- **Smart Report Detection**: Uses flexible targeting to find the best available reports (individual reports preferred, consolidated as fallback)
- **Organized File Management**: Downloads are systematically organized by company with consistent naming
- **Handles Real-World Complexity**: Different companies have different report types available - this tool adapts automatically
- **GitHub Actions Integration**: Automated quarterly downloads aligned with MOPS filing deadlines

## ✨ Key Features

- 📥 **Flexible Targeting System**: Prioritizes individual financial reports but falls back to consolidated reports when needed
- 📁 **Clean Organization**: Files saved in `downloads/{company_id}/` with standardized naming
- 🛡️ **Robust Error Handling**: Handles SSL issues, encoding problems, and missing reports gracefully  
- 📊 **Comprehensive Analysis**: Shows exactly what reports were found and why they were selected/rejected
- 🔄 **Two Operating Modes**: Flexible mode (default) for maximum success, strict mode for individual reports only
- 📝 **Detailed Logging**: Complete audit trail of all operations and decisions
- 🤖 **Automated Scheduling**: GitHub Actions workflow runs on MOPS filing deadlines with 5-day retry windows

## 🚀 Quick Start

### Installation

1. **Clone the repository**:
```bash
git clone https://github.com/your-repo/mops-downloader.git
cd mops-downloader
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

### Basic Usage

**Download all quarters for a company (recommended)**:
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024
```

**Download specific quarter**:
```bash
python scripts/mops_downloader.py --company_id 8272 --year 2023 --quarter 2
```

**Use strict mode (individual reports only)**:
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024 --strict_mode
```

### Batch Processing

**Update stock list and download all companies**:
```bash
# First, update the stock list
python Get觀察名單.py

# Then download reports for all companies in the list
python DownloadAll.py --year 2024 --quarter 1
```

## 📋 Input Parameters

| Parameter | Type | Description | Example | Default |
|-----------|------|-------------|---------|---------|
| `company_id` | String | Taiwan stock company ID | "2330", "8272" | Required |
| `year` | Integer | Reporting year (Western format) | 2024, 2023 | Required |
| `quarter` | Integer/String | Quarter (1-4) or "all" | 1, 2, 3, 4, "all" | "all" |
| `strict_mode` | Boolean | Only download individual reports | True/False | False |
| `output` | String | Output directory | "./reports" | "./downloads" |

## 🎯 Understanding Report Types

The system intelligently handles different types of financial reports:

### Primary Targets (Preferred)
- **IFRSs個別財報** - Individual Financial Reports (A12.pdf)
- **IFRSs個體財報** - Individual Financial Reports (A13.pdf)

### Secondary Targets (Fallback)
- **IFRSs合併財報** - Consolidated Financial Reports (AI1.pdf, A1L.pdf)
- **財務報告書** - General Financial Reports

### Always Excluded
- **英文版** - English versions
- **AIA.pdf**, **AE2.pdf** - English consolidated reports

## 📂 Output Structure

```
downloads/
├── 2330/                           # Company folder
│   ├── 202401_2330_AI1.pdf        # Q1 2024
│   ├── 202402_2330_AI1.pdf        # Q2 2024  
│   ├── 202403_2330_AI1.pdf        # Q3 2024
│   ├── 202404_2330_AI1.pdf        # Q4 2024
│   └── metadata.json              # Download metadata
├── 8272/
│   ├── 202401_8272_A12.pdf
│   └── metadata.json
└── logs/
    └── mops_downloader_20240805_143022.log
```

**File Naming**: `YYYYQQ_{company_id}_{report_type}.pdf`
- `YYYY`: Year (2024)
- `QQ`: Quarter (01, 02, 03, 04)
- `{company_id}`: Company stock ID
- `{report_type}`: A12, A13, AI1, etc.

## 💡 Usage Examples

### Example 1: Taiwan Semiconductor (TSMC) - Company 2330
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024
```

**Expected Result**: Downloads consolidated reports (AI1.pdf) as individual reports aren't available
```
✅ Downloaded: 202401_2330_AI1.pdf, 202402_2330_AI1.pdf, 202403_2330_AI1.pdf, 202404_2330_AI1.pdf
📊 Used consolidated reports as fallback (no individual reports available)
```

### Example 2: Systex Corporation - Company 8272
```bash
python scripts/mops_downloader.py --company_id 8272 --year 2024
```

**Expected Result**: Downloads individual reports (A12.pdf) - preferred type
```
✅ Downloaded: 202401_8272_A12.pdf, 202402_8272_A12.pdf, 202403_8272_A12.pdf, 202404_8272_A12.pdf
📊 Used individual reports (primary target achieved)
```

### Example 3: Mixed Availability - Company 2382
```bash
python scripts/mops_downloader.py --company_id 2382 --year 2023
```

**Expected Result**: Partial success with clear explanation
```
✅ Downloaded: 202304_2382_A13.pdf
❌ Missing: Q1, Q2, Q3 (only consolidated reports available, individual reports found for Q4 only)
```

## 🤖 Automated Downloads

### GitHub Actions Integration

The system includes automated quarterly downloads that run on a schedule aligned with Taiwan's MOPS filing deadlines.

#### MOPS Filing Schedule

Taiwan's Market Observation Post System (MOPS) requires companies to file quarterly reports by specific deadlines. Our automated system downloads reports immediately after these deadlines:

| Quarter | Period | Filing Deadline | Auto-Download Window |
|---------|--------|----------------|----------------------|
| **Q1** | Jan-Mar | **May 15** | May 15-19 (5 days) |
| **Q2** | Apr-Jun | **Aug 14** | Aug 14-18 (5 days) |
| **Q3** | Jul-Sep | **Nov 14** | Nov 14-18 (5 days) |
| **Q4** | Oct-Dec | **March 31** (next year) | March 31 - April 4 (5 days) |

#### How It Works

1. **Automatic Execution**: GitHub Actions runs on filing deadline dates at 02:00 UTC
2. **5-Day Retry Window**: Attempts download for 5 consecutive days to catch late filings
3. **Smart Skip Logic**: Only downloads missing files (won't re-download existing PDFs)
4. **Matrix Upload**: Automatically uploads status matrix to Google Sheets (if configured)
5. **Auto-Commit**: Commits all downloaded PDFs and metadata to the repository
6. **Comprehensive Logging**: Creates detailed logs and status reports for each run

#### Why Downloads Run AFTER Filing Deadlines

Reports are published **after** quarters end, so downloads are scheduled accordingly:

**Example: Q1 2025 Timeline**
```
├── Quarter Period: January 1 - March 31, 2025
├── Quarter Ends: March 31, 2025
├── Filing Deadline: May 15, 2025 ← Companies must file by this date
└── Auto-Download: May 15-19, 2025 ✅ Reports are now available!

Why the delay?
- Q1 doesn't end until March 31
- Companies need time to prepare financial statements
- Legal filing deadline is May 15 (45 days after quarter end)
- Most companies file near the deadline
- 5-day window ensures we catch all filings
```

**Example: Q4 2025 Timeline**
```
├── Quarter Period: October 1 - December 31, 2025
├── Quarter Ends: December 31, 2025
├── Filing Deadline: March 31, 2026 ← Next year!
└── Auto-Download: March 31 - April 4, 2026 ✅ Reports are now available!

Why March 2026?
- Q4 2025 is the annual report (full year)
- Companies get until March 31 of NEXT year to file
- This is 90 days after year-end for comprehensive audit
- Auto-download runs in March/April 2026 for Q4 2025 data
```

#### Manual Trigger

You can manually trigger downloads via GitHub Actions without waiting for the scheduled runs:

**Steps:**
1. Go to your repository on GitHub
2. Click the **"Actions"** tab
3. Select **"Download MOPS PDFs"** workflow from the left sidebar
4. Click **"Run workflow"** button (top right)
5. Configure parameters:
   - **Year**: Target year (e.g., 2025, 2024)
   - **Quarter**: Specific quarter (1, 2, 3, or 4)
   - **Delay**: Seconds between downloads (default: 10.0)
   - **Start from**: Optional company ID to start from (default: 2412)
   - **Skip existing files**: ✅ Recommended (only download missing files)
   - **Upload to sheets**: ✅ Enable for Google Sheets matrix view
6. Click **"Run workflow"** to start

**Use Cases for Manual Trigger:**
- Download historical data for past years
- Re-download specific quarters if needed
- Test the workflow with custom parameters
- Download immediately without waiting for scheduled run

#### Monitoring Downloads

**Check Download Status:**

- **Actions Tab**: View real-time workflow execution logs
  - See which companies are being processed
  - Track download progress and errors
  - View retry attempts (1/5, 2/5, etc.)

- **Commits**: Look for automated commit messages
  - `📥 Scheduled MOPS Download (Retry 1/5) - 2025 Q1`
  - `📥 Scheduled MOPS Download (Retry 2/5) - 2025 Q1`
  - Shows number of files downloaded

- **Google Sheets**: Matrix view (if configured)
  - Worksheet: "MOPS下載狀態"
  - Shows comprehensive download status for all companies
  - Updated automatically after each run

- **Repository Files**: Direct file inspection
  - Check `downloads/` folder for new PDFs
  - Review `logs/` for detailed execution logs
  - Check `data/reports/` for CSV matrix backups

**Example Commit Messages:**
```
📥 Scheduled MOPS Download (Retry 1/5) + 📊 Matrix Upload - 2025 Q1 (95 files from 110 companies)
📥 Scheduled MOPS Download (Retry 2/5) + 📊 Matrix Upload - 2025 Q1 (8 files from 12 companies)
📥 Scheduled MOPS Download (Retry 3/5) + 📊 Matrix Upload - 2025 Q1 (2 files from 3 companies)
```

## 🔧 Python API Usage

```python
from mops_downloader import MOPSDownloader

# Initialize downloader
downloader = MOPSDownloader(
    download_dir="./financial_reports",
    strict_mode=False,  # Use flexible targeting
    log_level="INFO"
)

# Download reports
result = downloader.download("2330", 2024, "all")

# Check results
if result.success:
    print(f"✅ Successfully downloaded {result.total_files} files")
    print(f"📁 Files: {result.downloaded_files}")
    print(f"💾 Total size: {result.total_size:,} bytes")
else:
    print(f"❌ Download failed: {result.error_details}")

# Handle partial success
if result.missing_quarters:
    print(f"⚠️ Missing quarters: {', '.join(result.missing_quarters)}")
```

## 📊 Understanding the Output

### Successful Download
```
[INFO] 📊 Report Analysis:
[INFO]    ✅ Target reports found: 4
[INFO]       • IFRSs個別財報 → 202401_8272_A12.pdf (Matched primary target)
[INFO]       • IFRSs個別財報 → 202402_8272_A12.pdf (Matched primary target)
[INFO]    📋 Consolidated reports: 0
[INFO]    🌍 English reports: 0
[INFO] ✅ Download completed successfully: 4/4 files (12.3 MB total)
```

### Partial Success with Explanation
```
[INFO] 📊 Report Analysis:
[INFO]    ✅ Target reports found: 1
[INFO]       • IFRSs個體財報 → 202304_2382_A13.pdf (Matched primary target)
[INFO]    📋 Consolidated reports: 3 (excluded in flexible mode preference)
[INFO] ⚠️ Download completed with missing quarters: 1/4 files
[INFO] ❌ Q1, Q2, Q3: No individual reports available
```

## 🛠️ Configuration

### Environment Setup
```bash
# Optional: Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
# or
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt
```

### Common Configuration Options
```python
# In your script or config file
DOWNLOAD_CONFIG = {
    'verify_ssl': False,           # Needed for MOPS compatibility
    'rate_limit_delay': 1.0,       # Seconds between requests
    'max_retries': 3,              # Retry attempts for failed downloads
    'timeout': 30,                 # Request timeout in seconds
    'strict_mode': False           # Use flexible targeting by default
}
```

### GitHub Actions Setup

To enable automated downloads, configure these repository secrets:

1. Go to **Settings** → **Secrets and variables** → **Actions**
2. Add the following secrets:
   - `GOOGLE_SHEETS_CREDENTIALS`: Your Google service account JSON (optional)
   - `GOOGLE_SHEET_ID`: Your Google Sheets spreadsheet ID (optional)

**Note**: Google Sheets integration is optional. The system will generate CSV backups even without Sheets credentials.

## 🔍 Troubleshooting

### Common Issues

**SSL Certificate Errors**:
```
Solution: SSL verification is automatically disabled for MOPS compatibility
```

**Encoding Issues**:
```
Solution: The system automatically handles Big5/UTF-8 encoding conversion
```

**No Reports Found**:
```
Check: 1) Company ID is correct 2) Year/quarter has data 3) Try flexible mode
```

**Partial Downloads**:
```
This is normal - not all companies have all report types for all quarters
Check the detailed log output for explanation
```

**GitHub Actions Not Running**:
```
Check:
1. Workflow file is in .github/workflows/Download.yaml
2. Actions are enabled in repository settings
3. Scheduled time hasn't arrived yet (check cron schedule)
```

### Debug Mode
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024 --log_level DEBUG
```

## 📁 Project Structure

```
mops-downloader/
├── mops_downloader/           # Main package
│   ├── downloads/             # Download management
│   ├── parsers/              # HTML/document parsing
│   ├── storage/              # File management
│   ├── validators/           # Input validation
│   └── web/                  # Web navigation
├── scripts/
│   └── mops_downloader.py    # Main CLI script
├── .github/workflows/
│   └── Download.yaml         # GitHub Actions automation
├── DownloadAll.py            # Batch download all companies
├── Get觀察名單.py             # Update stock list
├── StockID_TWSE_TPEX.csv    # Taiwan stock company list
├── downloads/                # Downloaded files (created automatically)
├── logs/                     # Log files (created automatically)
└── requirements.txt          # Python dependencies
```

## 📈 Requirements

- **Python**: 3.9 or higher
- **Dependencies**: See `requirements.txt`
- **Network**: Internet connection for MOPS access
- **Disk Space**: Varies by usage (PDFs are typically 1-5 MB each)
- **GitHub Actions**: Optional, for automated downloads

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Make your changes
4. Add tests if applicable
5. Commit your changes (`git commit -am 'Add new feature'`)
6. Push to the branch (`git push origin feature/new-feature`)
7. Create a Pull Request

## 📞 Support

- **Documentation**: See `instructions.md` for detailed technical specifications
- **Issues**: Report bugs or request features via GitHub issues
- **Logs**: Check `logs/` directory for detailed error information
- **Actions**: Monitor GitHub Actions tab for automated download status

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📝 Version History

### v2.0.0 (Current)
- ✅ Flexible targeting system with intelligent fallbacks
- ✅ Two-step download process for improved reliability  
- ✅ Comprehensive report analysis and categorization
- ✅ Enhanced error handling and logging
- ✅ Support for modern MOPS file patterns
- ✅ GitHub Actions automation with MOPS deadline alignment
- ✅ 5-day retry window for maximum success rate
- ✅ Google Sheets matrix integration

### v1.0.0
- Basic individual report downloading
- Simple file organization
- Core functionality

---

**Note**: This tool is designed to work with Taiwan's MOPS system and handles the complexities of real-world financial report availability. The flexible targeting system ensures maximum download success while providing clear explanations for any missing reports. Automated downloads run on Taiwan's official filing deadlines to ensure reports are available when downloaded.