require 'spec_helper'

describe 'My behaviour' do
  describe 'libreoffice' do
    it 'simple_text' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/editor_specific_documents/libreoffice/simple_text.xlsx')
      expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('This is a test')
    end
  end

  describe 'ms_office_2013' do
    it 'simple_text' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/editor_specific_documents/ms_office_2013/simple_text.xlsx')
      expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('This is a test')
    end
  end
end
