# Gas_for_Gmail_Distribution
## Gメールで複数名に複数種類の本文を配信するためのspread sheet
 - シンプルなシート構成  
 - 本文は任意の数のパターンを用意できます  
 
## How to Use  
1. 配信したいメールの件名をB1、本文をB2～に入力  
　複数パターンある場合はその分だけシートを設け、任意のシート名を設定する  
2. 「送信先」シートに宛名とメールアドレス、件名・本文として設定したいシートのシート名を入力
3. 署名シートを設定  
（手順1~3の実施は順不同）
  
4. 実行  
  - __生成内容をテストしたいとき__　34行目「GmailApp.sendEmail(～」をコメントアウト（文頭に//をつけてコメント化）し、  
   36行目「GmailApp.createDraft(～」をコメントアウト（文頭に//をつけてコメント化）  
  - __実際に生成して配信するとき__　コメント化されている36行目「GmailApp.sendEmail(～」をアンコメント（文頭//を削除）し、  
   36行目「GmailApp.createDraft(～」をコメントアウト（文頭に//をつけてコメント化）  
   
## 作成結果 
|  送信先  |  TD  |
|  件名  |  TD  |
|  本文 |  TD  |

## spread sheet画像 
![送信先シート_resized](https://user-images.githubusercontent.com/97077671/235069774-ec89996c-d95d-4b38-a4cf-51b129238658.png)  
![A_resized](https://user-images.githubusercontent.com/97077671/235069782-919b037c-abc6-4737-86bd-addf3ecf66f2.png)  
![B_resized](https://user-images.githubusercontent.com/97077671/235069785-de09c6f5-e8a2-4804-b171-e23cc2f98b80.png)  
![署名シート](https://user-images.githubusercontent.com/97077671/235069797-09d5a8e2-19d9-4292-b694-cc2cdea4e4ab.png)  
