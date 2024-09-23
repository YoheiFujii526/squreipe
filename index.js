const PORT = 8000;

const express = require("express");
const axios = require("axios");
const cheerio = require("cheerio");
const xlsx = require('xlsx');
const fs = require('fs');
const moment = require('moment');
const currentTime = moment();

let max_page;

const workbook = xlsx.utils.book_new();

const app = express();

//基準のページや終わりのページを決めてスクレイピングを開始する関数　
async function scrapeMultiplePages () {
    const baseUrl = 'https://www.amazon.co.jp/s?k=%E3%82%AD%E3%83%BC%E3%83%9C%E3%83%BC%E3%83%89&page=';//基本となるURL
    const startPage = 1;
    const endPage = max_page;

    for(let i=startPage;i<=endPage;i++) {

        const Url = `${baseUrl}${i}`;
        const scrapeData = [];
        console.log(`Scraping page ${i}: ${Url}`);

        await new Promise(resolve => setTimeout(resolve, 1000)); // 1秒の遅延を挿入
        console.log(`\n---${i}のデータを表示開始---\n`);
        await DoubleScrape(Url, scrapeData, i);
        
        console.log(`\n---${i}のデータを表示終了---\n`);
    }
}

//特定のクラスの要素をスクレイプする関数
async function DoubleScrape(url, data, page) {
    
    try {
        const response = await axios(url);
        const htmlParser = response.data;
        const $ = cheerio.load(htmlParser);

        $(".a-section.a-spacing-small.puis-padding-left-small.puis-padding-right-small").each(function() {
            const title = $(this).find(".a-size-base-plus").text().trim();
            const price = $(this).find(".a-price-whole").text().trim();
            const merchandiseUrl = "https://www.amazon.co.jp" + $(this).find("a").attr("href").trim();


            if (title && price && merchandiseUrl) {
                data.push({title, merchandiseUrl, price});
            }
        });

        console.log(data);  // データを表示する

        writeToExcel(data, page);
    } catch (error) {
        console.error(`Error fetching ${url}:`, error);
    }
}

// データをExcelファイルに書き込む関数
function writeToExcel(data, page) {
    const worksheetData = data.map(item => [item.title, item.merchandiseUrl, item.price]);
    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, `Sheet${page}`);
    // Excelファイルを保存
    const filename = `scraped_data${currentTime}.xlsx`;
    xlsx.writeFile(workbook, filename);
  
    console.log(`Data has been written to ${filename}`);
}



let URL = "https://www.amazon.co.jp/s?k=%E3%82%AD%E3%83%BC%E3%83%9C%E3%83%BC%E3%83%89";
axios(URL).then((response) => {
    let htmlParser = response.data;
    //console.log(htmlParser);
    let $ = cheerio.load(htmlParser);
    $(".s-pagination-strip", htmlParser).each(function() {
        let num = $(this).find(".s-pagination-item.s-pagination-disabled").text().trim();
        max_page = parseInt(num.replace(/[^0-9]/g, ''));
        console.log("maxpage=" + max_page);
    });

    scrapeMultiplePages();
    
}).catch((error) => console.log(error));


app.listen(PORT, console.log("server running!"));
