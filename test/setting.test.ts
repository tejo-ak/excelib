import {excelUtilTest} from '../src/excelutil'
const {ExcelUtil} =excelUtilTest
describe('calculate', function() {
    it('add', function() {
      let result = ExcelUtil.calcAddress(2,3,"A5");
      expect(result).toBe("C6");   
  })
})