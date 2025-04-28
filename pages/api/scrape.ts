import type { NextApiRequest, NextApiResponse } from 'next';
import axios from 'axios';
import ExcelJS from 'exceljs';
import https from 'https';
import cheerio from 'cheerio';

interface NewsItem {
  title: string;
  link: string;
  source: string;    // 언론사명
  reporter: string;  // 기자명
  published: string;
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  const clientId = process.env.NAVER_CLIENT_ID;
  const clientSecret = process.env.NAVER_CLIENT_SECRET;
  if (!clientId || !clientSecret) {
    return res
      .status(500)
      .json({ error: 'NAVER_CLIENT_ID 및 NAVER_CLIENT_SECRET을 .env.local에 설정하세요.' });
  }

  try {
    // SSL 인증서 오류 우회: 검증 비활성화
    const httpsAgent = new https.Agent({ rejectUnauthorized: false });
    const { data } = await axios.get('https://openapi.naver.com/v1/search/news.json', {
      httpsAgent,
      params: { query: 'GS칼텍스', display: 100, sort: 'date' },
      headers: {
        'X-Naver-Client-Id': clientId,
        'X-Naver-Client-Secret': clientSecret,
      },
    });

    const items: NewsItem[] = data.items.map((item: any) => ({
      title: item.title.replace(/<[^>]+>/g, ''),
      link: item.link,
      source: item.source || '',
      reporter: item.author || '',
      published: new Date(item.pubDate).toLocaleString('ko-KR'),
    }));

    const excludeKeywords = ['배구', '여자배구', '포토','리그','여자부','승점','세트','시즌','프로배구','프로'];
    const filteredItems = items.filter(item =>
      !excludeKeywords.some(keyword => item.title.includes(keyword))
    );

    // 언론사 한글명 매핑 테이블
    const pressMap: Record<string, string> = {
      "www.news.kbs.co.kr": "KBS",
      "www.imbc.com": "MBC",
      "www.sbs.co.kr": "SBS",
      "www.ebs.co.kr": "EBS",
      "ysmbc.co.kr":"여수MBC",
      "news.sbs.co.kr":"SBS",
      "www.jtbc.co.kr": "JTBC",
      "www.tvchosun.com": "TV조선",
      "www.channela.co.kr": "채널A",
      "www.mbn.co.kr": "MBN",
      "www.ichannela.com":"채널A",
      "yna.co.kr": "연합뉴스",
      "ytn.co.kr": "YTN",
      "www.ytn.co.kr":"YTN",
      "news.mtn.co.kr":"MTN뉴스",
      "news.dealsitetv.com":"딜사이트경제TV",
      "biz.sbs.co.kr":"SBS BIZ",
      "www.paxetv.com":"PAX경제TV",
      "www.chosun.com":"조선일보",
      "www.joongang.co.kr":"중앙일보",
      "www.donga.com":"동아일보",
      "www.hani.co.kr": "한겨레신문",
      "www.khan.co.kr": "경향신문",
      "www.kmib.co.kr": "국민일보",
      "www.seoul.co.kr": "서울신문",
      "www.asiatoday.co.kr":"아시아투데이",
      "www.hankookilbo.com":"한국일보",
      "www.segye.com":"세계일보",
      "www.naeil.com":"내일신문",
      "www.munhwa.com":"문화일보",
      "www.mk.co.kr": "매일경제신문",
      "www.hankyung.com": "한국경제신문",
      "www.sedaily.com": "서울경제신문",
      "www.fnnews.com": "파이낸셜뉴스",
      "www.asiae.co.kr": "아시아경제신문",
      "www.ajunews.com":"아주경제",
      "view.asiae.co.kr":"아시아경제",
      "www.thebell.co.kr":"더벨",
      "www.edaily.co.kr": "이데일리",
      "www.biztribune.co.kr": "비즈트리뷴",
      "news.mt.co.kr": "머니투데이",
      "biz.heraldcorp.com": "헤럴드경제",
      "www.businesspost.co.kr": "비지니스포스트",
      "news.bizwatch.co.kr": "비즈와치",
      "mydaily.co.kr": "마이데일리",
      "www.inews24.com": "아이뉴스24",
      "www.goodkyung.com": "굿모닝경제",
      "www.dailian.co.kr": "데일리안",
      "www.enewstoday.co.kr":"이뉴스투데이",
      "www.hansbiz.co.kr":"한스경제",
      "www.viva100.com":"브릿지경제",
      "www.womaneconomy.co.kr":"여성경제신문",
      "www.mediapen.com":"미디어펜",
      "www.moneys.co.kr":"머니S",
      "www.etoday.co.kr":"이투데이",
      "www.enetnews.co.kr":"이넷뉴스",
      "www.financialpost.co.kr":"파이낸셜포스트",
      "www.startuptoday.co.kr":"오늘경제",
      "www.srtimes.kr":"SR타임스",
      "biz.newdaily.co.kr":"뉴데일리경제",
      "www.segyebiz.com":"세계비즈",
      "www.seoulfn.com":"서울파이낸스",
      "biz.chosun.com":"조선비즈",
      "dealsite.co.kr":"딜사이트",
      "www.ezyeconomy.com":"이지경제",
      "www.greened.kr":"녹색경제신문",
      "www.econonews.co.kr":"이코노뉴스",
      "www.seouleconews.com":"서울이코노미뉴스",
      "www.newsprime.co.kr":"프라임경제",
      "www.widedaily.com":"와이드경제",
      "www.economytalk.kr":"이코노미톡뉴스",
      "www.finomy.com":"현대경제신문",
      "www.smartfn.co.kr":"스마트에프엔",
      "www.getnews.co.kr":"글로벌경제신문",
      "www.econovill.com":"이코노믹리뷰",
      "www.ebn.co.kr":"EBN산업경제",
      "www.kfenews.co.kr":"한국금융경제신문",
      "www.newsis.com": "뉴시스",
      "www.mt.co.kr": "머니투데이",
      "www.yna.co.kr": "연합뉴스",
      "www.newspim.com":"뉴스핌",
      "www.news1.kr":"뉴스원",
      "news.einfomax.co.kr":"연합인포맥스",
      "isplus.com": "일간스포츠",
      "www.osen.co.kr":"OSEN",
      "www.sportsseoul.com":"스포츠서울",
      "www.xportsnews.com":"엑스포츠뉴스",
      "sports.donga.com":"스포츠동아",
      "www.dailysportshankook.co.kr":"데일리스포츠한국",
      "www.mhnse.com":"MHN스포츠",
      "www.stoo.com":"스포츠투데이",
      "sports.chosun.com":"스포츠조선",
      "www.kjdaily.com":"광주매일신문",
      "www.kwangju.co.kr":"광주일보",
      "www.namdonews.com":"남도일보",
      "gj.newdaily.co.kr":"뉴데일리 호남제주",
      "www.kookje.co.kr":"국제신문",
      "www.mdilbo.com":"무등일보",
      "www.busan.com":"부산일보",
      "www.kyeonggi.com":"경기일보",
      "www.idaegu.co.kr":"대구신문",
      "www.jndn.com":"전남매일",
      "www.netongs.com":"여수넷통뉴스",
      "www.m-i.kr":"매일일보",
      "www.sisajournal.com": "시사저널",
      "amenews.kr":"신소재경제",
      "www.etnews.com":"전자신문",
      "www.cstimes.com":"컨슈머타임스",
      "it.chosun.com":"IT조선",
      "www.aflnews.co.kr":"농수축산신문",
      "kpenews.com":"한국정경신문",
      "zdnet.co.kr":"ZDNET",
      "www.healthinnews.co.kr":"헬스인뉴스",
      "www.00news.co.kr":"공공뉴스",
      "www.dt.co.kr":"디지털타임스",
      "www.techm.kr":"테크엠",
      "www.itdaily.kr":"IT데일리",
      "www.digitaltoday.co.kr":"디지털투데이",
      "www.lcnews.co.kr":"라이센스뉴스",
      "www.ekn.kr":"에너지경제",
      "www.nocutnews.co.kr":"노컷뉴스",
      "www.safetimes.co.kr": "세이프타임즈",
      "www.shinailbo.co.kr": "신아일보",
      "www.newsworker.co.kr":"뉴스워커",
      "www.siminsori.com":"시민의소리",
      "www.cnbnews.com":"CNB뉴스",
      "www.newscj.com" :"천지일보",
      "www.metroseoul.co.kr":"메트로",
      "www.ceoscoredaily.com":"CEO스코어데일리",
      "www.smarttimes.co.kr" :"스마트타임즈",
      "www.g-enews.com":"글로벌이코노믹",
      "www.fieldnews.kr":"필드뉴스",
      "www.newsway.co.kr":"뉴스웨이",
      "www.newsworks.co.kr":"뉴스웍스",
      "www.ziksir.com":"직썰",
      "www.smarttoday.co.kr":"스마트투데이",
      "www.topdaily.kr":"톱데일리",
      "www.youthdaily.co.kr":"청년일보",
      "www.straightnews.co.kr":"스트레이트뉴스",
      "www.breaknews.com":"브레이크뉴스",
      "www.whitepaper.co.kr":"화이트페이퍼",
      "www.thebigdata.co.kr":"빅데이터뉴스",
      "m.skyedaily.com":"스카이데일리",
      "www.newsian.co.kr":"뉴시안",
      "www.thefirstmedia.net":"더퍼스트미디어",
      "www.sisaon.co.kr":"시사오늘",
      "news.tf.co.kr" : "더팩트",
      "www.asiaa.co.kr":"아시아A",
      "www.kpinews.kr":"KPI뉴스",
      "www.koreaittimes.com":"코리아IT타임스",
      "www.popcornnews.net":"팝콘뉴스",
      "www.opinionnews.co.kr":"오피니언뉴스",
      "www.newsdream.kr":"뉴스드림",
      "www.newsclaim.co.kr":"뉴스클레임",
      "daily.hankooki.com":"데일리한국",
      "www.theguru.co.kr":"더구루",
      "www.maniareport.com":"마니아타임즈",
      "www.ngetnews.com":"뉴스저널리즘",
      "www.newsquest.co.kr":"뉴스퀘스트",
      "www.topstarnews.net":"톱스타뉴스",
      "www.fetv.co.kr":"FETV",
      "www.meconomynews.com":"메경이코노미",
      "www.bizwnews.com":"비즈월드"
    };
    // originallink 기반 도메인 추출하여 source 설정
    filteredItems.forEach(item => {
      try {
        const domain = new URL((item as any).originallink || item.link).hostname;
        item.source = pressMap[domain] || domain;
      } catch {
        item.source = '';
      }
    });

    // 각 기사 페이지에서 기자명 추출
    for (const item of filteredItems) {
      try {
        const resp = await axios.get(item.link, { httpsAgent });
        const $ = cheerio.load(resp.data);
        const metaAuthor = $('meta[name="author"]').attr('content');
        const byline = $('.byline, .article_writer').first().text();
        item.reporter = (metaAuthor || byline || '').trim();
      } catch (e) {
        // 스크래핑 실패 시 빈 문자열
        item.reporter = '';
      }
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('GS칼텍스 뉴스');
    const finalItems = filteredItems;
    sheet.columns = [
      { header: '발행일', key: 'published', width: 20 },
      { header: '제목', key: 'title', width: 50 },
      { header: '언론사', key: 'source', width: 30 },
      { header: '기자명', key: 'reporter', width: 20 },
      { header: '링크', key: 'link', width: 50 },
    ];
    finalItems.forEach((row) => sheet.addRow(row));

    const buffer = await workbook.xlsx.writeBuffer();
    // 한글 파일명 지원: RFC5987 `filename*` 형식 사용
    const fileName = 'GS칼텍스뉴스.xlsx';
    res.setHeader(
      'Content-Disposition',
      `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.send(buffer);
  } catch (err: any) {
    console.error('API error:', err);
    const msg = err.message || String(err);
    res.status(500).json({ error: msg });
  }
}
