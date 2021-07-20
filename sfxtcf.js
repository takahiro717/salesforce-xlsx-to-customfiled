/*
使い方
$ node sfxtcf.js ./samples/JP_CustomFieldTest__c.xlsx
引数に対象のエクセルを入力する。
※テキストエリアは文字数の入力があるとエラーになる
*/

/* --- 初期設定 --- */
/* ------------------------------------------------------------------------- */
// xlsx エクセル読み込み
const XLSX = require("xlsx");
const workbook = XLSX.readFile(process.argv[2]); // 引数でエクセルを指定 ※process.argv[2]=第一引数
const sheet = workbook.Sheets["Field"];

// jsforce メタデータの保存と更新
const jsforce = require('jsforce');
const conn = new jsforce.Connection({ loginUrl: 'https://login.salesforce.com/' }); //ログインURLの指定が必要な場合
const username = ""; //ログイン用ユーザーネーム
const password = "";// パスワードとセキュリティトークン スペース無しでつなげる IP制限を解除しているとトークンは不要
const excelCol = 300; //13以上の数値、エクセル行の300まで確認する。それ以上の場合は数値を変更する ※自動取得が安定しないらしいので固定値にした

// 項目シートの列定義 ※エクセルの列の増減をやりやすくするためのオブジェクト
const sheetCol = {
    label: "A", //ラベル（label）
    fullName: "B", //項目名
    type: "C", //データ型
    required: "D", //必須
    unique: "E", //一意
    externalId: "F", //外部ID
    inlineHelpText: "G", //ヘルプテキスト
    length: "H", //文字数
    visibleLines: "I", //行数
    precision: "J", //桁数
    scale: "K", //小数点
    valueSetName: "L", //グローバル選択リスト値
    valueSet_valueSetDefinition_value: "M", //選択リスト値
    valueSetDefault: "N", //選択リスト先頭行デフォルト設定
    valueSet_valueSetDefinition_sorted: "O", //選択リストアルファベット順ソート
    valueSet_restricted: "P", //値セットで定義された値に選択リストを制限します
    valueSet_controllingField: "Q", //制御項目
    valueSet_valueSettings: "R", //項目の連動関係
    formulaType: "S", //数式：戻り値
    formula: "T", //数式
    formulaTreatBlanksAs: "U", //数式：空白時
    displayFormat: "V", //自動採番：表示形式
    startingNumber: "W", //自動採番：開始番号
    summaryForeignKey: "X", //積み上げ集計：対象オブジェクト
    summaryOperation: "Y", //積み上げ集計：種別
    summarizedField: "Z", //積み上げ集計：積み上げ項目
    summaryFilterItems: "AA", //積み上げ集計：検索条件
    referenceTo: "AB", //主従と参照関係
    relationshipLabel: "AC", //主従と参照関係：項目の表示ラベル
    relationshipName: "AD", //主従と参照関係：子リレーション名
    lookupFilter: "AE", //主従と参照関係：ルックアップ検索条件
    deleteConstraint: "AF", //参照関係：参照レコードが削除された場合の対処方法
    reparentableMasterDetail: "AG", //主従関係：親の変更を許可
    writeRequiresMasterRead: "AH", //主従関係：共有設定
    defaultValueCheckBox: "AI", //初期値：チェックボックス
    defaultValueFormula: "AJ", //初期値：数値・テキスト
    description: "AK", //説明
    profileName1: "AL5", //カスタム項目セキュリティ１の名前
    profile1: "AL", //カスタム項目セキュリティ １
    profileName2: "AM5", //カスタム項目セキュリティ２の名前
    profile2: "AM", //カスタム項目セキュリティ２
    profileName3: "AN5", //カスタム項目セキュリティ３の名前
    profile3: "AN", //カスタム項目セキュリティ３
    profileName4: "AO5", //カスタム項目セキュリティ４の名前
    profile4: "AO", //カスタム項目セキュリティ４
    profileName5: "AP5", //カスタム項目セキュリティ５の名前
    profile5: "AP", //カスタム項目セキュリティ５
    profileName6: "AQ5", //カスタム項目セキュリティ６の名前
    profile6: "AQ", //カスタム項目セキュリティ６
    profileName7: "AR5", //カスタム項目セキュリティ７の名前
    profile7: "AR", //カスタム項目セキュリティ７
    profileName8: "AS5", //カスタム項目セキュリティ８の名前
    profile8: "AS", //カスタム項目セキュリティ８
    profileName9: "AT5", //カスタム項目セキュリティ９の名前
    profile9: "AT", //カスタム項目セキュリティ９
    profileName10: "AU5", //カスタム項目セキュリティ１０の名前
    profile10: "AU", //カスタム項目セキュリティ１０
    FIX: "AV", //FIX
}

// カスタム項目用配列の宣言（中身はオブジェクトの配列）
const customFields = getCustomFieldsFromXslx();

// フィールドレベルセキュリティ設定用プロファイル定義配列 Admin=システム管理者
const profiles = getProfilesFromXslx(); // エクセルからデータを登録する際にはこちらの行を有効にする

// フィールドレベルセキュリティ用オブジェクト配列の宣言
const fieldPermissions = getPermissionsFromXslx(profiles);

// console用 色指定
const consoleColorRed = '\u001b[31m';
const consoleColorReset = '\u001b[0m';

// パース用
// const xml2js = require("xml2js");

//デバッグ用ファイル書き出し設定
//const fs = require("fs"); // ログ等のファイル書き込み用
//const out = fs.createWriteStream('info.log'); // 標準出力をリダイレクト
//const infoLog = new console.Console(out);

//デバッグ用 ログ
//console.log(customFields);
//infoLog.log(customFields);
//infoLog.log(customFields[13].valueSet); //リスト値が入っているか確認するときに使う
//infoLog.log(customFields[13].valueSet.valueSetDefinition.value[0]); //リスト値が入っているか確認するときに使う
//infoLog.log(customFields[13].valueSet.valueSetDefinition.value[1]); //リスト値が入っているか確認するときに使う
//infoLog.log(customFields[13].valueSet.valueSetDefinition.value[2]); //リスト値が入っているか確認するときに使う
//infoLog.log(profiles);
//infoLog.log(fieldPermissions);
//infoLog.log(fieldPermissions[0].profilePermisson);

//デバッグ用 JSON形式でワークブック全体の内容をファイル出力されるので、成形して閲覧する
/*
fs.writeFile('workbook.txt', JSON.stringify(workbook), (err, data) => {
    if (err) console.log(err);
    else console.log('write end');
});
*/

/**
 * ログインしてからUPSERTと権限登録の処理をする
 */
conn.login(username, password)
    .then(() => {
        // API制限回避のため、配列を10個ずつに分割して処理している
        upsert();
    }, err => {
        console.error(err);
    });

/**
 * JSforceでカスタム項目のUPSERTをする
 * @param {object[]} slicedFields 10件ずつのカスタムフィールドメタデータオブジェクトの配列
 */
async function upsert() {
    for (let i = 0; i < customFields.length; i += 10) {
        let slicedFields = customFields.slice(i, i + 10);
        await conn.metadata.upsert('CustomField', slicedFields)
            .then(results => {
                // 結果が1件のときは配列ではなくオブジェクトで返ってくる
                if (Array.isArray(results) == false) {
                    upsertResultDisplay(results);
                } else {
                    for (let result of results) {
                        upsertResultDisplay(result);
                    }
                }
            }, err => {
                if (err) { console.error(err); }
            });
    }
    security();
}

/**
 * JSforceでカスタム項目のUPSERTをした後の結果表示
 * @param {object} result Jsforceから返ってきた結果
 */
function upsertResultDisplay(result) {
    if (result.success == false) {
        console.log(consoleColorRed + 'upsert result : ' + result.success + ' : ' + result.fullName + consoleColorReset);
    } else {
        console.log('upsert result : ' + result.success + ' : ' + result.fullName);
    }
}

/**
 * JSforceでカスタム項目レベルセキュリティの設定をする
 * @param {string[]} profiles プロファイル名の配列
 * @param {object[]} fieldPermissions 
 */
function security() {
    for (let i = 0; i < profiles.length; i++) { //プロファイルの数だけforで処理
        conn.metadata.update('Profile', { fullName: profiles[i], fieldPermissions: fieldPermissions[i].profilePermisson })
            .then(results => {
                if (results.success == false) {
                    console.log(consoleColorRed + 'set permission result : ' + results.success + ' : ' + results.fullName + consoleColorReset);
                } else {
                    console.log('set permission result : ' + results.success + ' : ' + results.fullName);
                }
            }, err => {
                console.error(err);
            });
    }
}


/* --- カスタム項目登録用の情報をエクセルから取得 --- */
/* ------------------------------------------------------------------------- */
/* 
// info.logに

// valueSetDefinitionの中身
エクセルに以下のように書かれている場合のサンプル
----
福岡
東京
----
console.log(customFields[0].valueSet.valueSetDefinition.value[1]);
{ default: false, fullName: '東京', label: '東京' }
*/

/**
 * カスタム項目登録用の情報をエクセルから取得
 * @returns {object[]} カスタムフィールドメタデータオブジェクトの配列
 */
function getCustomFieldsFromXslx() {
    let fields = [];
    let cnt = 0; //配列用カウンター
    let value = ""; //エクセルの文字から置き換えが必要なときに利用する変数
    for (let i = 7; i <= excelCol; i++) { //i はエクセルの項目の開始行
        if (sheet[sheetCol.label + i] && !sheet[sheetCol.FIX + i]) { //A列にデータが存在しておりFIXではないか確認

            // 配列にオブジェクトを追加
            fields.push({});

            // ラベル（label）
            fields[cnt].label = sheet[sheetCol.label + i]['v'];

            // 項目名（fullName）
            // エクセルA3からオブジェクト名を取得して連結する
            fields[cnt].fullName = sheet['A3']['v'] + "." + sheet[sheetCol.fullName + i]['v'];

            // データ型（type）
            // チェックボックスのデフォルトはfalseで固定
            value = sheet[sheetCol.type + i]['v'];
            if (value == "自動採番" || value == "Auto Number") { value = "AutoNumber"; }
            else if (value == "積み上げ集計" || value == "Roll-Up Summary") { value = "Summary"; }
            else if (value == "外部参照関係" || value == "External Lookup Relationship") { value = "ExternalLookup"; }
            else if (value == "参照関係" || value == "Lookup Relationship") { value = "Lookup"; }
            else if (value == "主従関係" || value == "Master-Detail Relationship") { value = "MasterDetail"; }
            else if (value == "URL") { value = "Url"; }
            else if (value == "チェックボックス" || value == "Checkbox") { value = "Checkbox"; fields[cnt].defaultValue = false; }
            else if (value == "テキスト" || value == "Text") { value = "Text"; }
            else if (value == "テキスト(暗号化) " || value == "Text (Encrypted)") { value = "EncryptedText"; }
            else if (value == "テキストエリア" || value == "Text Area") { value = "TextArea"; }
            else if (value == "パーセント" || value == "Percent") { value = "Percent"; }
            else if (value == "メール" || value == "Email") { value = "Email"; }
            else if (value == "テキストエリア (リッチ)" || value == "Text Area (Rich)") { value = "Html"; }
            else if (value == "ロングテキストエリア" || value == "Text Area (Long)") { value = "LongTextArea"; }
            else if (value == "数値" || value == "Number") { value = "Number"; }
            else if (value == "選択リスト" || value == "Picklist") { value = "Picklist"; }
            else if (value == "選択リスト (複数選択)" || value == "Picklist (Multi-Select)") { value = "MultiselectPicklist"; }
            else if (value == "地理位置情報" || value == "Geolocation") { value = "Location"; }
            else if (value == "通貨" || value == "Currency") { value = "Currency"; }
            else if (value == "電話" || value == "Phone") { value = "Phone"; }
            else if (value == "日付" || value == "Date") { value = "Date"; }
            else if (value == "日付/時間" || value == "Date/Time") { value = "DateTime"; }
            fields[cnt].type = value;
            /*
            自動採番	Auto Number
            数式	Formula
            積み上げ集計	Roll-Up Summary
            係外部参照関	External Lookup Relationship
            参照関係	Lookup Relationship
            主従関係	Master-Detail Relationship
            URL	URL
            チェックボックス	Checkbox
            テキスト	Text
            テキスト(暗号化) 	Text (Encrypted)
            テキストエリア	Text Area
            パーセント	Percent
            メール	Email
            テキストエリア (リッチ)	Text Area (Rich)
            ロングテキストエリア	Text Area (Long)
            数値	Number
            選択リスト	Picklist
            選択リスト (複数選択)	Picklist (Multi-Select)
            地理位置情報	Geolocation
            通貨	Currency
            電話	Phone
            日付	Date
            日付/時間	Date/Time
            */

            // 必須（required）
            if (sheet[sheetCol.required + i]) {
                fields[cnt].required = true;
            }

            // 一意（unique）
            if (sheet[sheetCol.unique + i]) {
                fields[cnt].unique = true;
            }

            // 外部ID（externalId）
            if (sheet[sheetCol.externalId + i]) {
                fields[cnt].externalId = true;
            }

            // 選択リスト：グローバル選択リスト値セット（valueSet.valueSetName）
            if (sheet[sheetCol.valueSetName + i]) {
                fields[cnt].valueSet = { valueSetName: sheet[sheetCol.valueSetName + i]['v'] };
            }

            // 選択リスト：値を指定（valueSet.valueSetDefinition）
            if (sheet[sheetCol.valueSet_valueSetDefinition_value + i]) {
                let valueSetDefault = false;
                fields[cnt].valueSet = { valueSetDefinition: { value: [] } };
                let listdata = sheet[sheetCol.valueSet_valueSetDefinition_value + i]['v'].split("\r\n"); //\r\nでsplit
                for (let j = 0; j < listdata.length; j++) {
                    if (j == 0) { // ループの初回のみデフォルト設定の判定を行う
                        if (sheet[sheetCol.valueSetDefault + i]) { valueSetDefault = true; } //先頭行デフォルト設定 記述があればtrue
                    } else {
                        valueSetDefault = false; //ループ2週目以降はfalseを設定する
                    }
                    fields[cnt].valueSet.valueSetDefinition.value[j] = {
                        default: valueSetDefault,
                        fullName: listdata[j],
                        label: listdata[j]
                    };
                }
            }

            // 選択リスト：アルファベット順のソート（valueSet.valueSetDefinition.sorted）
            if (sheet[sheetCol.valueSet_valueSetDefinition_sorted + i]) {
                fields[cnt].valueSet.valueSetDefinition.sorted = true;
            }

            // 選択リスト：制御項目（valueSet.restricted）
            if (sheet[sheetCol.valueSet_restricted + i]) {
                fields[cnt].valueSet.restricted = true;
            }

            // 選択リスト：制御項目（valueSet.controllingField）
            if (sheet[sheetCol.valueSet_controllingField + i]) {
                fields[cnt].valueSet.controllingField = sheet[sheetCol.valueSet_controllingField + i]['v'];
            }

            // 選択リスト：項目の連動関係（valueSet.valueSettings
            if (sheet[sheetCol.valueSet_valueSettings + i]) {
                fields[cnt].valueSet.valueSettings = JSON.parse(sheet[sheetCol.valueSet_valueSettings + i]['v']).valueSettings;
            }

            // 文字数（length）
            if (sheet[sheetCol.length + i]) {
                fields[cnt].length = Number(sheet[sheetCol.length + i]['v']);
            }

            // 行数（visibleLines）
            // 選択リスト (複数選択)のときは３以上 ロングテキストは２以上
            if (sheet[sheetCol.visibleLines + i]) {
                fields[cnt].visibleLines = Number(sheet[sheetCol.visibleLines + i]['v']);
            }

            // 桁数（precision） 256.99 = 5
            if (sheet[sheetCol.precision + i]) {
                fields[cnt].precision = Number(sheet[sheetCol.precision + i]['v']);
            }

            // 小数点（scale） 256.99 = 2
            if (sheet[sheetCol.scale + i]) {
                fields[cnt].scale = Number(sheet[sheetCol.scale + i]['v']);
            }

            // 数式：戻り値（type）※数式の戻り値をtypeに設定する
            // 数値、パーセント、通貨は桁数と小数点が必須になる
            if (sheet[sheetCol.formulaType + i]) {
                value = sheet[sheetCol.formulaType + i]['v'];
                if (value == "チェックボックス" || value == "Checkbox") { value = "Checkbox"; }
                else if (value == "テキスト" || value == "Text") { value = "Text"; }
                else if (value == "数値" || value == "Number") { value = "Number"; }
                else if (value == "パーセント" || value == "Percent") { value = "Percent"; }
                else if (value == "通貨" || value == "Currency") { value = "Currency"; }
                else if (value == "日付" || value == "Date") { value = "Date"; }
                else if (value == "日付/時間" || value == "Date/Time") { value = "DateTime"; }
                fields[cnt].type = value;
            }
            /*
            チェックボックス	Checkbox
            テキスト	Text
            数値	Number
            パーセント	Percent
            通貨	Currency
            日付	Date
            日付/時間	Date/Time
            */

            // 数式（formula）
            if (sheet[sheetCol.formula + i]) {
                fields[cnt].formula = sheet[sheetCol.formula + i]['v'];
            }

            // 数式：空白時（formulaTreatBlanksAs）
            // リファレンスが「formulaTreatBlankAs」になっていた
            if (sheet[sheetCol.formulaTreatBlanksAs + i]) {
                if (sheet[sheetCol.formulaTreatBlanksAs + i]['v'] == "空白" || sheet[sheetCol.formulaTreatBlanksAs + i]['v'] == "Blank") {
                    fields[cnt].formulaTreatBlanksAs = 'BlankAsBlank'; // 「空白」がエクセルに入力されている場合
                } else {
                    fields[cnt].formulaTreatBlanksAs = 'BlankAsZero'; // 「0」がエクセルに入力されている場合
                }
            }
            /*
            空白 Blank
            */

            // 自動採番：表示形式（displayFormat） A-{00000} = A-00100
            if (sheet[sheetCol.displayFormat + i]) {
                fields[cnt].displayFormat = sheet[sheetCol.displayFormat + i]['v'];
            }

            // 自動採番：開始番号（startingNumber） 100 = 100からスタート
            if (sheet[sheetCol.startingNumber + i]) {
                fields[cnt].startingNumber = Number(sheet[sheetCol.startingNumber + i]['v']);
            }

            // 積み上げ集計：対象オブジェクト（summaryForeignKey）
            if (sheet[sheetCol.summaryForeignKey + i]) {
                fields[cnt].summaryForeignKey = sheet[sheetCol.summaryForeignKey + i]['v'];
            }

            // 積み上げ集計：種別（summaryOperation）
            if (sheet[sheetCol.summaryOperation + i]) {
                if (sheet[sheetCol.summaryOperation + i]['v'] == '件数' || sheet[sheetCol.summaryOperation + i]['v'] == 'COUNT') { fields[cnt].summaryOperation = 'Count'; }
                else if (sheet[sheetCol.summaryOperation + i]['v'] == '合計' || sheet[sheetCol.summaryOperation + i]['v'] == 'SUM') { fields[cnt].summaryOperation = 'Sum'; }
                else if (sheet[sheetCol.summaryOperation + i]['v'] == '最大' || sheet[sheetCol.summaryOperation + i]['v'] == 'MAX') { fields[cnt].summaryOperation = 'Max'; }
                else if (sheet[sheetCol.summaryOperation + i]['v'] == '最小' || sheet[sheetCol.summaryOperation + i]['v'] == 'MIN') { fields[cnt].summaryOperation = 'Min'; }
            }
            /*
            件数  COUNT
            合計  SUM
            最大  MAX
            最少  MIN
            */

            // 積み上げ集計：積み上げ項目（summarizedField）
            if (sheet[sheetCol.summarizedField + i]) {
                fields[cnt].summarizedField = sheet[sheetCol.summarizedField + i]['v'];
            }

            // 積み上げ集計：検索条件（summaryFilterItems）
            if (sheet[sheetCol.summaryFilterItems + i]) {
                fields[cnt].summaryFilterItems = JSON.parse(sheet[sheetCol.summaryFilterItems + i]['v']).summaryFilterItems;
            }

            // 主従と参照関係（referenceTo）
            if (sheet[sheetCol.referenceTo + i]) {
                fields[cnt].referenceTo = sheet[sheetCol.referenceTo + i]['v'];
                // 参照関係ラベルの自動指定
                fields[cnt].relationshipLabel = sheet['B3']['v']; //オブジェクトラベル エクセルのB3に書く
                //子リレーション名の自動指定
                let fromName = sheet['A3']['v'].replace('__c', ''); //子リレーション名用に__cを削除
                let toName = sheet[sheetCol.referenceTo + i]['v'].replace('__c', ''); // 子リレーション名用に__cを削除
                fields[cnt].relationshipName = "Relation_" + fromName + "_to_" + toName;
                //参照関係のときに削除オプションを設定する
                if (sheet[sheetCol.type + i]['v'] == "参照関係" || sheet[sheetCol.type + i]['v'] == "Lookup Relationship") {
                    if (sheet[sheetCol.required + i]) { //必須項目の場合「参照関係に含まれる参照レコードは削除できません。」に設定
                        fields[cnt].deleteConstraint = "Restrict";
                    } else { //必須項目ではない場合、「この項目の値をクリアします。 この項目を必須にした場合、このオプションは選択できません。」に設定
                        fields[cnt].deleteConstraint = "SetNull";
                    }
                }
            }

            // 主従と参照関係：ラベル（relationshipLabel）
            if (sheet[sheetCol.relationshipLabel + i]) {
                fields[cnt].relationshipLabel = sheet[sheetCol.relationshipLabel + i]['v'];
            }

            // 主従と参照関係：子リレーション名（relationshipName）
            // 指定がある場合にここで上書きを行う
            if (sheet[sheetCol.relationshipName + i]) {
                fields[cnt].relationshipName = sheet[sheetCol.relationshipName + i]['v'];
            }

            // 主従と参照関係：ルックアップ検索条件（lookupFilter）
            if (sheet[sheetCol.lookupFilter + i]) {
                fields[cnt].lookupFilter = JSON.parse(sheet[sheetCol.lookupFilter + i]['v']).lookupFilter;
            }

            // 参照関係：参照レコードが削除された場合の対処方法（deleteConstraint）
            if (sheet[sheetCol.deleteConstraint + i]) {
                fields[cnt].deleteConstraint = sheet[sheetCol.deleteConstraint + i]['v'];
            }

            // 主従関係：親の変更を許可（reparentableMasterDetail）
            if (sheet[sheetCol.reparentableMasterDetail + i]) {
                fields[cnt].reparentableMasterDetail = true;
            }

            // 主従関係：共有設定（writeRequiresMasterRead）
            if (sheet[sheetCol.writeRequiresMasterRead + i]) {
                fields[cnt].writeRequiresMasterRead = true;
            }

            // デフォルト値：チェックボックス（defaultValue）
            if (sheet[sheetCol.defaultValueCheckBox + i]) {
                fields[cnt].defaultValue = true;
            }

            // デフォルト値：テキスト（defaultValue）
            if (sheet[sheetCol.defaultValueFormula + i]) {
                fields[cnt].defaultValue = sheet[sheetCol.defaultValueFormula + i]['v'];
            }

            // ヘルプテキスト（inlineHelpText）
            if (sheet[sheetCol.inlineHelpText + i]) {
                fields[cnt].inlineHelpText = sheet[sheetCol.inlineHelpText + i]['v'];
            }

            // 説明（description）
            if (sheet[sheetCol.description + i]) {
                fields[cnt].description = sheet[sheetCol.description + i]['v'];
            }

            cnt++; //次のループ用に配列カウンターに+1をする
        }
    }
    return fields;
}

/* --- カスタム項目権限の登録（エクセルシートから） --- */
/* ------------------------------------------------------------------------- */
/*
profiles :
[ 'Admin', 'Partner Community User' ]
fieldPermissions :
[
  {
    profilePermisson: [
      [Object], [Object],
      [Object], [Object],
      [Object], [Object],
      [Object], [Object]
    ]
  },
  {
    profilePermisson: [
      [Object], [Object],
      [Object], [Object],
      [Object], [Object],
      [Object], [Object]
    ]
  }
]
*/

/**
 * エクセルのプロファイル名を配列に入れる
 * @returns {string[]} プロファイル名の配列
 */
function getProfilesFromXslx() {
    let profiles = [];
    if (sheet[sheetCol.profileName1]) { profiles.push(sheet[sheetCol.profileName1]['v']) }
    if (sheet[sheetCol.profileName2]) { profiles.push(sheet[sheetCol.profileName2]['v']) }
    if (sheet[sheetCol.profileName3]) { profiles.push(sheet[sheetCol.profileName3]['v']) }
    if (sheet[sheetCol.profileName4]) { profiles.push(sheet[sheetCol.profileName4]['v']) }
    if (sheet[sheetCol.profileName5]) { profiles.push(sheet[sheetCol.profileName5]['v']) }
    if (sheet[sheetCol.profileName6]) { profiles.push(sheet[sheetCol.profileName6]['v']) }
    if (sheet[sheetCol.profileName7]) { profiles.push(sheet[sheetCol.profileName7]['v']) }
    if (sheet[sheetCol.profileName8]) { profiles.push(sheet[sheetCol.profileName8]['v']) }
    if (sheet[sheetCol.profileName9]) { profiles.push(sheet[sheetCol.profileName9]['v']) }
    if (sheet[sheetCol.profileName10]) { profiles.push(sheet[sheetCol.profileName10]['v']) }
    return profiles;
}

/**
 * カスタム項目レベルセキュリティ設定をエクセルから取得
 * @param {string[]} profiles プロファイル名の配列
 * @returns {object[]} プロファイルの数だけオブジェクトが入ったもの
 */
function getPermissionsFromXslx(profiles) {
    let permissions = []; // プロファイルの数だけオブジェクトを格納する配列
    let set = []; // 各条件の真偽値を格納する配列
    // プロファイルの数だけループ
    for (let j = 0; j < profiles.length; j++) {
        let cnt2 = 0; //
        permissions.push({ profilePermisson: [] }); // プロファイル毎にの中にカスタム項目セキュリティ設定を入れる配列を作る
        // プロファイル毎のカスタム項目セキュリティ設定を入れる配列にjsforce用のオブジェクトを入れていく
        for (let i = 7; i <= excelCol; i++) {
            //主従関係、数式、必須項目は処理から外す
            if (sheet[sheetCol.label + i] != null
                && sheet[sheetCol.type + i]['v'] != "主従関係"
                && sheet[sheetCol.type + i]['v'] != "Master-Detail Relationship"
                && sheet[sheetCol.required + i] == null) {

                permissions[j].profilePermisson.push({}); // 行単位のオブジェクトを追加

                // カスタム項目（field）
                permissions[j].profilePermisson[cnt2].field = sheet['A3']['v'] + "." + sheet[sheetCol.fullName + i]['v'];

                // editableとreadableを関数から取得 if文の中に書くと長くなるので関数化した
                if (j == 0) { set = selectPermission(sheet[sheetCol.profile1 + i]['v']); }
                if (j == 1) { set = selectPermission(sheet[sheetCol.profile2 + i]['v']); }
                if (j == 2) { set = selectPermission(sheet[sheetCol.profile3 + i]['v']); }
                if (j == 3) { set = selectPermission(sheet[sheetCol.profile4 + i]['v']); }
                if (j == 4) { set = selectPermission(sheet[sheetCol.profile5 + i]['v']); }
                if (j == 5) { set = selectPermission(sheet[sheetCol.profile6 + i]['v']); }
                if (j == 6) { set = selectPermission(sheet[sheetCol.profile7 + i]['v']); }
                if (j == 7) { set = selectPermission(sheet[sheetCol.profile8 + i]['v']); }
                if (j == 8) { set = selectPermission(sheet[sheetCol.profile9 + i]['v']); }
                if (j == 9) { set = selectPermission(sheet[sheetCol.profile10 + i]['v']); }

                // 編集権限（editable）
                permissions[j].profilePermisson[cnt2].editable = set[0];

                // 参照権限（readable）
                permissions[j].profilePermisson[cnt2].readable = set[1];

                cnt2++;
            }
        }
    }
    return permissions;
}

/**
 * カスタム項目セキュリティの条件分岐
 * @param {string} value 編集 or 参照 or 閲覧不可
 * @returns {string[]} editable、readableの順番で返す
 */
function selectPermission(value) {
    if (value == '編集' || value == 'Edit') {
        return ['true', 'true']
    }
    if (value == '参照' || value == 'Readonly') {
        return ['false', 'true']
    }
    if (value == '閲覧不可' || value == 'None') {
        return ['false', 'false']
    }
}