require 'spec_helper'

describe OoxmlParser::DocProperties do
  it 'doc_properties title' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/doc_properties/doc_properties.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content.office2010_content.doc_properties.title).to eq('Shape Title')
  end

  it 'doc_properties description' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/doc_properties/doc_properties.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content.office2010_content.doc_properties.description).to eq('Shape Description')
  end
end
