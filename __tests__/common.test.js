/*
モックのjsonファイルを取得する

*/
const forceClass = require('../src/common.js');
const force = new forceClass();

test('test', () => {
    expect(force.selectPermission("編集")).toStrictEqual(['true', 'true']);
    expect(force.selectPermission("参照")).toStrictEqual(['false', 'true']);
    expect(force.selectPermission("閲覧不可")).toStrictEqual(['false', 'false']);
    expect(force.selectPermission("Edit")).toStrictEqual(['true', 'true']);
    expect(force.selectPermission("Readonly")).toStrictEqual(['false', 'true']);
    expect(force.selectPermission("None")).toStrictEqual(['false', 'false']);
});