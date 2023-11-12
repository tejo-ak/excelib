//require('../src/excelutil.ts')
import '../dist/excelutil.js'
describe('calculate', function() {
    it('add', function() {
      let result = ExcelUtil.calcAddress(2,3,"A5");
      expect(result).toBe("C6");   
  })
})