require 'spec_helper'

describe 'My behaviour' do
  it 'text_field.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/text_field/text_field.xlsx')
    expect(xlsx.worksheets[0].drawings.first.grouping.elements[1].text_body.paragraphs.first.text_field).to be_a(OoxmlParser::TextField)
  end
end
