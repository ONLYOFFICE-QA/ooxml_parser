require 'spec_helper'

describe 'My behaviour' do
  it 'table_layout_autofit' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_layout/table_layout_autofit.docx')
    expect(docx.elements[1].table_properties.table_layout.type).to eq(:autofit)
  end
end
