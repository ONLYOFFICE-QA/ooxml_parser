require 'spec_helper'

describe 'My behaviour' do
  it 'table_style_column_band_size_value' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_style_column_band_size/table_style_column_band_size_value.docx')
    expect(docx.elements[1].properties.table_style_column_band_size.value).to eq(5)
  end
end
