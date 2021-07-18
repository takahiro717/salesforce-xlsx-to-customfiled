import { Core } from '../core.js';
const core = new Core();


test('test', () => {
    expect(core.selectPermission("編集")).toStrictEqual(['true', 'true']);
    expect(core.selectPermission("参照")).toStrictEqual(['false', 'true']);
    expect(core.selectPermission("閲覧不可")).toStrictEqual(['false', 'false']);
    expect(core.selectPermission("Edit")).toStrictEqual(['true', 'true']);
    expect(core.selectPermission("Readonly")).toStrictEqual(['false', 'true']);
    expect(core.selectPermission("None")).toStrictEqual(['false', 'false']);
});