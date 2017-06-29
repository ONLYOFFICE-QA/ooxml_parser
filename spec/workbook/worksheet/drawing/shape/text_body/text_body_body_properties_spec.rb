require 'spec_helper'

describe 'Text Body Body Properties' do
  it 'number_columns' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/body_properties/body_properties_number_columns.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.properties.number_columns).to eq(5)
  end

  it 'column_distance' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/body_properties/body_properties_column_distance.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.properties.space_columns).to eq(OoxmlParser::OoxmlSize.new(2, :centimeter))
  end
end
