# frozen_string_literal: true

require 'spec_helper'

describe 'ParagraphProperties#paragraph_style_ref' do
  let(:docx) do
    OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties' \
                              '/paragraph_style_ref/paragraph_style_ref.docx')
  end

  let(:style_ref) { docx.elements.first.paragraph_properties.paragraph_style_ref }

  it 'paragraph_style_ref is correct class' do
    expect(style_ref).to be_a(OoxmlParser::ParagraphStyleRef)
  end

  it 'paragraph_style_ref has value' do
    expect(style_ref.value).to eq('693')
  end

  it 'paragraph_style_ref can find referenced style' do
    expect(style_ref.referenced_style).to be_a(OoxmlParser::DocumentStyle)
  end

  it 'referenced style has correct shade' do
    expect(style_ref.referenced_style.paragraph_properties.shade.fill)
      .to eq(OoxmlParser::Color.new(0, 255, 0))
  end
end
