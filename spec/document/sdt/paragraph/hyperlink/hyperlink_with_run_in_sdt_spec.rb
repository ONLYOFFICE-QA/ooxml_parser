# frozen_string_literal: true

require 'spec_helper'

describe 'sdt - paragraph - hyperlink - runs' do
  it 'hyperlink_with_runs.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/paragraph/hyperlink/runs/hyperlink_with_runs.docx')
    expect(docx.elements.first.sdt_content.paragraphs.first.hyperlink.runs.first.text).to eq('Heading1 - Page 2')
  end

  it 'run_with_space_preserve.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/paragraph/hyperlink/runs/run_with_space_preserve.docx')
    expect(docx.elements.first.sdt_content.paragraphs.first.hyperlink.runs[1].t.space).to eq(:preserve)
  end

  it 'hyperlink_run_with_tab.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/paragraph/hyperlink/runs/hyperlink_run_with_tab.docx')
    expect(docx.elements.first.sdt_content.paragraphs.first.hyperlink.runs[2].tab).to be_a(OoxmlParser::Tab)
  end
end
