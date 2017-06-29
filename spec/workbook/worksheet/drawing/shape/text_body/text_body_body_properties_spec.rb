require 'spec_helper'

describe 'Text Body Body Properties' do
  it 'number_columns' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/body_properties/body_properties_number_columns.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.properties.number_columns).to eq(5)
  end
end
