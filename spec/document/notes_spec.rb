require 'spec_helper'

describe OoxmlParser::Note do
  it 'HeaderFooterCount' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/notes/header_footer_count.docx')
    expect(docx.notes.length).to eq(4)
  end

  it 'Footer' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/notes/footer.docx')
    expect(docx.notes.first.type).to eq('footer1')
    expect(docx.notes.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end
end
