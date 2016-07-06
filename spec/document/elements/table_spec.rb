require 'spec_helper'

describe OoxmlParser::Table do
  it 'TableInTable' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/table_in_table.docx')
    expect(docx.elements[1].rows[0].cells[0].elements[1].is_a?(OoxmlParser::Table)).to eq(true)
  end
end
