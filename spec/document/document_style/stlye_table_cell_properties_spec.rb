require 'spec_helper'

describe 'My behaviour' do
  it 'table_cell_properties_border' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_cell_properties/table_cell_properties_border.docx')
    expect(docx.elements[1].properties.table_style.table_cell_properties.table_cell_borders.bottom.space).to eq(OoxmlParser::OoxmlSize.new(0))
  end
end
