require 'spec_helper'

describe 'My behaviour' do
  it 'fill_without_outline.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/spacing/run_properties_spacing.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs[1].properties.space).to eq(OoxmlParser::OoxmlSize.new(80, :twip))
  end
end
