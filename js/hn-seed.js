/* 호남지역단 교육생 초기 데이터 시드
 * 원본: BUILTIN_STUDENTS (고객컨설팅 MASTER과정 면담일지 시스템)
 * 43명 — 광주비전센터 20명 + 동광주비전센터 23명
 *
 * 필드 매핑:
 *   base(기준실적)   = 평균실적(원)
 *   target(목표실적) = 순증목표(원)
 *   honors(아너스실적) = 1월 + 2월 실적 합계(원)
 */
window.HN_SEED_STUDENTS = [
  // ===== 광주비전센터 =====
  { region: "호남지역단", center: "광주비전센터", branch: "광주TC지점",    empNo: "107681", name: "박희자", phone: "010-9965-2965", base: 756885,  target: 860000,  honors: 1513770 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주TC지점",    empNo: "1A2807", name: "문채원", phone: "010-3111-6647", base: 1334085, target: 1430000, honors: 2668170 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주TC지점",    empNo: "1B9374", name: "천수정", phone: "010-7759-4860", base: 498955,  target: 600000,  honors: 997910 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주TC지점",    empNo: "1A4011", name: "김윤영", phone: "010-3637-6670", base: 462755,  target: 560000,  honors: 925510 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주미래지점",  empNo: "967424", name: "최미선", phone: "010-5610-2537", base: 505685,  target: 610000,  honors: 1011370 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주지점",      empNo: "1A1080", name: "모상민", phone: "010-8667-1271", base: 1032215, target: 1130000, honors: 2064430 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주지점",      empNo: "153239", name: "서영란", phone: "010-5206-6546", base: 800805,  target: 900000,  honors: 1601610 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주지점",      empNo: "1B0395", name: "이두성", phone: "010-5099-6221", base: 767160,  target: 870000,  honors: 1534320 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주지점",      empNo: "1C3909", name: "허용순", phone: "010-3756-6660", base: 891155,  target: 990000,  honors: 1782310 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주지점",      empNo: "1C8703", name: "심은영", phone: "010-2906-2520", base: 446495,  target: 550000,  honors: 892990 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주직할지점",  empNo: "158408", name: "이진혁", phone: "010-9406-8385", base: 1789565, target: 1890000, honors: 3579130 },
  { region: "호남지역단", center: "광주비전센터", branch: "광주직할지점",  empNo: "1C2459", name: "김미영", phone: "010-8777-5115", base: 536370,  target: 640000,  honors: 0 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A0446", name: "김정희", phone: "010-8600-0390", base: 844855,  target: 940000,  honors: 1689710 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A6499", name: "유선은", phone: "010-8600-0390", base: 1009390, target: 1110000, honors: 2018780 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A6607", name: "문효숙", phone: "010-2796-4620", base: 689265,  target: 790000,  honors: 1378530 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A6961", name: "최길례", phone: "010-5554-8406", base: 331000,  target: 430000,  honors: 662000 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A0445", name: "강지향", phone: "010-6610-8534", base: 524105,  target: 620000,  honors: 1048210 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "9A0590", name: "정인숙", phone: "010-9066-7048", base: 260390,  target: 360000,  honors: 520780 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "1B1166", name: "박은주", phone: "010-8249-9907", base: 381280,  target: 480000,  honors: 762560 },
  { region: "호남지역단", center: "광주비전센터", branch: "신광주지점",    empNo: "1C7250", name: "박진욱", phone: "010-2621-9320", base: 608130,  target: 710000,  honors: 1216260 },

  // ===== 동광주비전센터 =====
  { region: "호남지역단", center: "동광주비전센터", branch: "광주제일지점", empNo: "986037", name: "정경화", phone: "010-5607-5627", base: 1429830, target: 1530000, honors: 2859660 },
  { region: "호남지역단", center: "동광주비전센터", branch: "광주제일지점", empNo: "9A1520", name: "신경애", phone: "010-3635-6970", base: 2037730, target: 2140000, honors: 4075460 },
  { region: "호남지역단", center: "동광주비전센터", branch: "광주제일지점", empNo: "9A1766", name: "김승미", phone: "010-3624-0370", base: 2589935, target: 2690000, honors: 5179870 },
  { region: "호남지역단", center: "동광주비전센터", branch: "광주제일지점", empNo: "1B1312", name: "윤정화", phone: "010-3501-3567", base: 565025,  target: 670000,  honors: 1130050 },
  { region: "호남지역단", center: "동광주비전센터", branch: "동광주TC지점", empNo: "129731", name: "서혜경", phone: "010-6605-6693", base: 761600,  target: 860000,  honors: 1523200 },
  { region: "호남지역단", center: "동광주비전센터", branch: "동광주TC지점", empNo: "1B8497", name: "류민희", phone: "010-7192-4027", base: 418000,  target: 520000,  honors: 836000 },
  { region: "호남지역단", center: "동광주비전센터", branch: "동광주TC지점", empNo: "1A2040", name: "정명순", phone: "010-4156-9994", base: 461882,  target: 560000,  honors: 923763 },
  { region: "호남지역단", center: "동광주비전센터", branch: "동광주TC지점", empNo: "9A0450", name: "갈재명", phone: "010-4251-0821", base: 608845,  target: 710000,  honors: 1217690 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "144470", name: "김인자", phone: "010-6602-6542", base: 365230,  target: 470000,  honors: 730460 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "154822", name: "송정량", phone: "010-8607-0112", base: 332155,  target: 430000,  honors: 664310 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "949966", name: "조은아", phone: "010-2681-7826", base: 1339105, target: 1440000, honors: 2678210 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "137553", name: "윤진희", phone: "010-8829-1045", base: 1071490, target: 1170000, honors: 2142980 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "984627", name: "박점희", phone: "010-4603-2502", base: 684030,  target: 780000,  honors: 1368060 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무중앙지점", empNo: "965034", name: "유해순", phone: "010-6659-1220", base: 535920,  target: 640000,  honors: 1071840 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무제일지점", empNo: "190467", name: "송지연", phone: "010-9213-0129", base: 438375,  target: 540000,  honors: 876750 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무제일지점", empNo: "1A0856", name: "강성원", phone: "010-9635-0191", base: 485055,  target: 590000,  honors: 970110 },
  { region: "호남지역단", center: "동광주비전센터", branch: "상무제일지점", empNo: "1D0562", name: "이종열", phone: "010-9455-8899", base: 286091,  target: 390000,  honors: 572182 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "132458", name: "박순자", phone: "010-3485-2510", base: 582965,  target: 680000,  honors: 1165930 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "979377", name: "나대수", phone: "010-2650-2700", base: 919045,  target: 1020000, honors: 1838090 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "9A0460", name: "강경희", phone: "010-3616-8895", base: 404655,  target: 500000,  honors: 809310 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "1C0027", name: "이병훈", phone: "010-3642-6794", base: 1043370, target: 1140000, honors: 2086740 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "1C6888", name: "김애덕", phone: "010-2916-6424", base: 498705,  target: 600000,  honors: 997410 },
  { region: "호남지역단", center: "동광주비전센터", branch: "새광주지점",   empNo: "127322", name: "김성자", phone: "010-9030-5209", base: 708550,  target: 810000,  honors: 1417100 }
];
