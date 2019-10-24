# frozen_string_literal: true

require 'spec_helper'

describe 'FontScheme' do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/theme/font_scheme/simple_docx.docx') }

  it 'Font scheme name is not empty' do
    expect(docx.theme.font_scheme.name).to eq('Office Classic 2')
  end

  it 'Major font has latin name' do
    expect(docx.theme.font_scheme.major_font.latin.typeface).to eq('Arial')
  end

  it 'Major font has east asian name' do
    expect(docx.theme.font_scheme.minor_font.east_asian.typeface).to eq('Arial')
  end

  it 'Major font has complex script name' do
    expect(docx.theme.font_scheme.minor_font.complex_script.typeface).to eq('Arial')
  end

  it 'Minor font has latin name' do
    expect(docx.theme.font_scheme.minor_font.latin.typeface).to eq('Arial')
  end
end
