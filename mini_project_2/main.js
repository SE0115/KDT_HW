// 상품 데이터
const data = [
    { name: '초콜렛', price: 2000 },
    { name: '아이스크림', price: 1000 },
    { name: '컵라면', price: 1600 },
    { name: '볼펜', price: 2500 },
    { name: '아메리카노', price: 4000 },
    { name: '과자', price: 3000 },
    { name: '탄산수', price: 1200 },
    { name: '떡볶이', price: 3500 },
    { name: '노트', price: 1500 },
    { name: '껌', price: 500 }
];

// 사용자 입력 받기
const line = prompt('최대 금액을 입력해주세요.');
const amount = parseInt(line);

// 주어진 금액으로 살 수 있는 가장 비싼 상품을 구함
const item = getItemByAmount(data, amount);

const msg = item ? 
    `${amount}원으로 살 수 있는 가장 비싼 상품은 [${item.name}]이고, 가격은 ${item.price}원입니다.` : 
    '살 수 있는 상품이 없습니다.';

// 결과 출력
alert(msg);

// 아래에 getItemByAmount 함수를 작성하세요.
function getItemByAmount(data, amount) {
    let item =null;
    let left_Money = amount;
    for(let i=0;i<data.length;i++){                                                           
        // 주어진 금액보다 낮은 가격의 물품과의 가격 차이를 min 값과 비교해 주어진 금액에 제일 근접한 상품 찾기
        if(amount-data[i].price >= 0 && amount-data[i].price < left_Money){ 
            left_Money = amount-data[i].price;
            item = data[i];
        }
    }
    return item;
}