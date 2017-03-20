# XLSX Converter ver.1

다음과 같은 순서로 진행하시면 됩니다.

1. Load the PriceByBarcode data
2. Select the Template
3. Select a XLSX file to convert
4. Click to download

PriceByBarcode 데이터는 아래의 양식을 가집니다.

| Barcode       | Name                                        | Price |
|:--------------|:-------------------------------------------:|:-----:|
| 8806164102947 | 美妆 韩国 得鲜The Saem 都市生态亚麻籽保湿精华 | 99    |

구현된 Template 들은 아래와 같습니다.

- 美俏
- C-MATE
- GET
- 达令
- 孩子王

각각의 Template을 로드하면 아래 `Target` 양식에 맞춰 컨버팅이 시작됩니다.

| Target      | 美俏 | C-MATE | GET | 达令 | 孩子王 |
| ----------- | :--: | :----: | :-: | :--: | :----: |
| 外部订单编号 |  A  |  A  |  A  |  A  |  A  |
| 商品条码 |  I  |  :x: |  :x:  |  J  |  N  |
| 快递公司 |  :x:  |  O  |  :x:  |  :x:  |  :x:  |
| 收件人名字 |  C  |  I  |  G  |  C  |  C  |
| 收件人身份证号 |  :x:  |  :x:  |  D  |  D  |  D  |
| 收件人省 |  :x:  |  (J)  |  (H)  |  (E)  |  E  |
| 收件人市 |  :x:  |  (J)  |  (H)  |  (E)  |  F  |
| 收件人县区 |  :x:  |  (J)  |  (H)  |  (E)  |  G  |
| 收件人街道及门牌号 |  D  |  K  |  (H)  |  (E)  |  H  |
| 收件人联系电话 |  E  |  M  |  I  |  F  |  J  |
| 邮件地址 |  :x:  |  N  |  (H)  |  :x:  |  :x:  |
| 包裹重量（KG） |  :x:  |  :x:  |  :x:  |  :x:  |  S  |
| 物品单价（RMB） |  (I)  |  :x:  |  :x:  |  (J)  |  (N)  |
| 物品数量 |  K  |  U  |  K  |  L  |  T  |
| 支付方式 |  :x:  |  :x:  |  :x:  |  :x:  |  :x:  |
| 商品名称 |  J  |  R  |  J  |  K  |  P  |
| 订单生成时间（年-月-日 时:分:秒） |  :x:  |  F  |  :x:  |  :x:  |  :x:  |
| 购买人平台ID |  :x:  |  :x:  |  :x:  |  :x:  |  :x:  |
| RACK NUMBER |  B  |  P  |  C  |  B  |  B  |

## 구현사항

### Web Workers Compatibility

Web Workers 미사용이 기본값이며, 활성화 시 Firefox에서만 동작합니다.
```html
<input type="checkbox" name="useworker" unchecked> <!-- if checked, it works in Firefox only!! -->
```

### Worksheet row_length 처리

```js
var row_length = workseheet['!ref'].substr(4);
```

- Worksheet는 'A1-S14'의 range를 가진다고 가정하고, `.substr(4)`하면 row_length 값인 14가 반환
- end에 해당하는 column 값이 Z를 넘어가면 동작하지 않음

### Update Worksheet range

When writing a worksheet by hand, be sure to update the range. For a longer discussion, see http://git.io/KIaNKQ

```js
worksheet['!ref'] = "A1:S"+row_length;
```

### 주소 문자열 처리

주소가 省(province), 市(city), 区(county), 나머지 상세주소(street and house number) 순서로 입력된다면 이를 구분하여 처리

```js
var string = "江苏省 无锡市 南长区 五星家园C区728号301室(214021)";
var strArray = string.split(" ");

/* province */
var strProvice = strArray[0];

/* city */
var strCity = strArray[1];

/* county */
var strCounty = strArray[2];

/* street and house number */
var restAddr = strArray[3];

if (strArray.length > 4) {
  for (var i=3; i<strArray.length; i++) {
    restAddr = restAddr.concat(" "+strArray[i]);
  }
}
```

- `美俏` template은 구현되지 않음(예외 주소 문자열)
- `GET` template은 mail address까지 처리

```js
restAddr = restAddr.substring(0, restAddr.length-8);
var mailAddr = restAddr.substring(restAddr.length-7, restAddr.length-1);
```

### Barcode

```js
function getPriceByBarcode(json, brcd) {
  return json.filter(function(item){
    return item.barcode.replace(/\s/g, '') == brcd; 
  });
}
```

### Writing Workbooks

```js
function work_cell(worksheet, address, value) {
  var w_cell;
  if (worksheet[address] == null) {
    worksheet[address] = {t:"s",v:"",w:""};
    w_cell = worksheet[address];
    w_cell.v = value;
    w_cell.w = value;
  } else {
    w_cell = worksheet[address];
    w_cell.t = "s";
    w_cell.v = value;
    w_cell.w = value;
  }
}
```

- raw value로 xlsx에 저장함

### Display converted data in HTML

```js
function htmlOut(data) {
  var output = JSON.stringify(to_json(data), 2, 2);
  if (out.innerText === undefined) {
    out.textContent = output;
  } else {
    out.innerText = output;
  } 
  if (typeof console !== 'undefined') {
    console.log("output", new Date());
  } 
}
```

### Download

```js
function download_xlsx() {
  var xout = XLSX.write(gWbOut, {bookType:'xlsx', bookSST:true, type: 'binary'});
  function s2ab_blob(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  var fooFileName = "cvtd." + gFnOut;
  saveAs(new Blob([s2ab_blob(xout)],{type:"application/octet-stream"}), fooFileName)
}
```

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of
the Microsoft Open Specifications Promise, falling under the same terms as
OpenOffice (which is governed by the Apache License v2).  Given the vagaries of
the promise, the Original Author makes no legal claim that in fact end users are
protected from future actions.  It is highly recommended that, for commercial
uses, you consult a lawyer before proceeding.