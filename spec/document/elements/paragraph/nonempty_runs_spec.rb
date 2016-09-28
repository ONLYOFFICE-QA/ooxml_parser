require 'spec_helper'

describe 'nonempty_runs' do
  it 'endnote_nonempty_runs' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/nonempty_runs/endnote_nonempty_runs.docx')
    expect(docx.elements.first.nonempty_runs.length).to eq(2)
  end

  it 'footnote_nonempty_runs' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/nonempty_runs/footnote_nonempty_runs.docx')
    expect(docx.elements.first.nonempty_runs.length).to eq(2)
  end
end
