require 'spec_helper'

describe 'Bookmark' do
  it 'bookmark_basic' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/bookmark/bookmark_basic.docx')
    expect(docx.elements.first.elements.first.name).to eq('Test')
  end

  it 'bookmark mark as nonempty run' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/bookmark/bookmark_basic.docx')
    expect(docx.elements.first.nonempty_runs.first.name).to eq('Test')
  end
end
