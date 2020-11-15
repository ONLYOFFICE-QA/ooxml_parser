# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Paragraph, '#align' do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/align/all_four_alignes.docx') }

  it 'Default align is symbol :left' do
    expect(docx.elements[0].align).to eq(:left)
  end

  it 'Center align is parsed correctly' do
    expect(docx.elements[1].align).to eq(:center)
  end

  it 'Right align is parsed correctly' do
    expect(docx.elements[2].align).to eq(:right)
  end

  it 'Justification is parsed correctly' do
    expect(docx.elements[3].align).to eq(:justify)
  end
end
