# frozen_string_literal: true

require 'spec_helper'

describe 'DocxShape#style' do
  files_dir = 'spec/document/elements/paragraph/' \
              'runs/alternate_content/office2010_content/' \
              'graphic/data/style'

  it 'style is nil for docx without style' do
    docx = OoxmlParser::Parser.parse("#{files_dir}/no_style.docx")
    expect(docx.element_by_description[0].character_style_array[1]
               .alternate_content
               .office2010_content.graphic.data
               .style).to be_nil
  end

  describe 'Parse not empty style' do
    docx = OoxmlParser::Parser.parse("#{files_dir}/style.docx")
    style = docx.element_by_description[0].character_style_array[1]
                .alternate_content
                .office2010_content.graphic.data.style

    it 'style have attribute for docx with style' do
      expect(style).not_to be_nil
    end

    it 'style has effect reference' do
      expect(style.effect_reference.index).to eq(0)
    end

    it 'style has fill reference' do
      expect(style.fill_reference.index).to eq(0)
    end

    it 'style has font reference' do
      expect(style.font_reference.index).to eq('minor')
    end

    it 'style has line reference' do
      expect(style.line_reference.index).to eq(0)
    end
  end
end
