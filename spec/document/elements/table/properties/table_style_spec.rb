require 'spec_helper'

describe 'Table Style' do
  it 'TableStyle root object' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_style/table_style_root_object.docx')
    expect(docx).to be_with_data
  end
end
