import puppeteer from 'puppeteer';
import XLSX from 'xlsx'

interface ResultItem {
  title: string
  href: string
  language: string
  product: string
  platform: string
}

interface ExcelObject {
  question: string;
  question_type: string;
  language: string;
  answer_url: string;
  answer_type: string;
  answer_content: string;
  product: string;
  platform: string;
  is_public: string;
  error_code: string;
  _ignore: string;
  question_index: string;
}

const LinkList = ['https://docs.agora.io/cn/Voice/product_voice?platform=Web', 'https://docs.agora.io/cn/Video/landing-page?platform=Web']

const initBrowser = async () => {
  const browser = await puppeteer.launch({
    defaultViewport: {
      width: 1920,
      height: 980
    }
  });
  const page = await browser.newPage();

  return {
    browser,
    page
  }
}

const getAllLinks = async (page: puppeteer.Page, url: string) => {
  await page.goto(url, { waitUntil: 'networkidle2' });

  const result = await page.evaluate(({url}) => {
    const sideBar = document.querySelector('.sidebar-menu')
    if(!sideBar) {
      return {
        error: `Error: 未找到链接目录，url: ${url}`
      }
    }

    const linkList = Array.from(sideBar.querySelectorAll('a'))
    return linkList.map(item => item.href)
  }, { url })

  if((result as any).error) {
    console.log((result as any).error)
    return []
  }

  return result
}

const getAllLinksFromList = async (page: puppeteer.Page, linkList: string[]) => {
  const result: string[] = []
  for(const link of linkList) {
    const links: any = await getAllLinks(page, link)
    result.push(...links)
  }

  return result
}

const getLinkResult = async (page: puppeteer.Page, url: string) => {
  await page.goto(url, { waitUntil: 'networkidle2' });

  const result = await page.evaluate(() => {
    const articleTitle: any = document.querySelector('.page-title')
    const container = document.querySelector('.article-page-container')
    if(!articleTitle || !container) {
      return []
    }

    const ret: ResultItem[] = []
    const url = new URL(window.location.href)
    const ProductMap = new Map([
      [
        'Voice',
        {
          en: 'Audio Call',
          cn: '语音通话'
        }
      ],
      [
        'Video', 
        {
          en: 'Video Call',
          cn: '视频通话'
        }
      ]
    ])

    let relativeH2Title = ''
    Array.from(container!.querySelectorAll('h2, h3')).map((item) => {
      let title: string = ''
      const product = ProductMap.get(url.pathname.split('/')[2]) || { cn: '未知产品', en: '' }

      if(item.tagName === 'H2') {
        relativeH2Title = item.id
        title = `${articleTitle.innerText} - ${item.id} - ${product.cn}`
      }
      if(item.tagName === 'H3') {
        title = `${articleTitle.innerText} - ${relativeH2Title} - ${item.id} - ${product.cn}`
      }

      url.hash = `#${item.id}`
      ret.push({
        title,
        href: url.href,
        language: url.pathname.split('/')[1],
        product: product.en,
        platform: url.searchParams.get('platform') || ''
      })
    })

    return ret
  });

  return result
}

const export2Excel = (result: ResultItem[]) => {
  const _headers: Array<keyof ExcelObject> = ['question', 'question_type', 'language', 'answer_url', 'answer_type', 'answer_content', 'product', 'platform', 'is_public', 'error_code', '_ignore', 'question_index']
  const _data: ExcelObject[] = result.map((item) => {
    return {
      question: item.title,
      question_type: '',
      language: item.language,
      answer_url: `<a href="${item.href}">${item.title}</a>`,
      answer_type: '文档链接',
      answer_content: '/',
      product: item.product,
      platform: item.platform,
      is_public: '',
      error_code: '',
      _ignore: '',
      question_index: ''
    }
  })

  const headers = _headers
    .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
    .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});

  const data = _data
    .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65 + j) + (i + 2) })))
    .reduce((prev, next) => prev.concat(next))
    .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});

  const output = Object.assign({}, headers, data);
  const outputPos = Object.keys(output);
  const ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];

  const wb = {
    SheetNames: ['mySheet'],
    Sheets: {
      'mySheet': Object.assign({}, output, { '!ref': ref }, {
        "!cols": [
          { wch: 60 },
          { wch: 15 },
          { wch: 10 },
          { wch: 120 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
          { wch: 15 },
        ]
      })
    }
  };

  XLSX.writeFile(wb, 'output.xlsx');
}

const main = async () => {
  const { browser, page } = await initBrowser()
  const linkList = await getAllLinksFromList(page, LinkList)
  const length = linkList.length
  console.log(`链接列表抓取完成，共${length}条。开始获取链接内容...`)

  const result: ResultItem[] = []
  for(const [index, link] of linkList.entries()) {
    const ret = await getLinkResult(page, link)
    result.push(...ret)

    const status = index + 1 === length ? '抓取完成!' : `正在抓取：${linkList[index + 1]}`
    console.log(`总进度: ${index + 1} / ${length}。${status}`)
  }

  await browser.close();
  export2Excel(result)
}

main()
