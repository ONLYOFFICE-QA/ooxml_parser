require 'spec_helper'

describe 'TableCellMargins' do
  it 'table_cell_margins_size' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_cell_margin/table_cell_margins_size.docx')
    expect(docx.elements[1].properties.table_style.table_properties.table_cell_margin.top).to eq(OoxmlParser::OoxmlSize.new(1.27, :centimeter))
  end
end
