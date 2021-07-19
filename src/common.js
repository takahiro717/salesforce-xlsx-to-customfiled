/*
Todo:
- 単体テストできるようにする
- モック準備
- セルの定義は環境変数にする
- Jsforceはテストしない
- XLSXはテストしない
- オブジェクトや配列作る部分をテストする
*/

module.exports = class Common {
    /**
     * エクセルのプロファイル名を配列に入れる
     * @returns {string[]} プロファイル名の配列
     */
    getProfilesFromXslx(sheet, sheetCol) {
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
    getPermissionsFromXslx(sheet, sheetCol, excelCol, profiles) {
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
                    if (j == 0 && sheet[sheetCol.profile1 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile1 + i]['v']); }
                    if (j == 1 && sheet[sheetCol.profile2 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile2 + i]['v']); }
                    if (j == 2 && sheet[sheetCol.profile3 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile3 + i]['v']); }
                    if (j == 3 && sheet[sheetCol.profile4 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile4 + i]['v']); }
                    if (j == 4 && sheet[sheetCol.profile5 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile5 + i]['v']); }
                    if (j == 5 && sheet[sheetCol.profile6 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile6 + i]['v']); }
                    if (j == 6 && sheet[sheetCol.profile7 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile7 + i]['v']); }
                    if (j == 7 && sheet[sheetCol.profile8 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile8 + i]['v']); }
                    if (j == 8 && sheet[sheetCol.profile9 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile9 + i]['v']); }
                    if (j == 9 && sheet[sheetCol.profile10 + i] != null) { set = this.selectPermission(sheet[sheetCol.profile10 + i]['v']); }

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
    selectPermission(value) {
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
}