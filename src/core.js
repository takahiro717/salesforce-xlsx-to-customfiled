export class Core {
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