# frozen_string_literal: true

require 'spec_helper'

describe 'Text Art' do
  it 'text_borders_width_1pt.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/text_borders_width_1pt.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to eq(OoxmlParser::OoxmlSize.new(1, :point))
  end

  it 'text_borders_width_2.25pt.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/text_borders_width_2.25pt.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to eq(OoxmlParser::OoxmlSize.new(2.25, :point))
  end

  it 'text_outline_color_standart.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/text_outline_color_standart.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(OoxmlParser::Color.new(255, 0, 0))
  end

  it 'text_outline_color_scheme.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/text_outline_color_scheme.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(OoxmlParser::Color.new(244, 177, 131))
  end

  it 'transform_circle.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/transform_circle.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.body_properties.preset_text_warp.preset).to eq(:textCircle)
  end

  it 'Transform_arc.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/tranform_arc.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.body_properties.preset_text_warp.preset).to eq(:textArchUp)
  end

  it 'opacity_50.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/opacity_50.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color.alpha_channel).to eq(50)
  end

  it 'opactiy_0.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/opactiy_0.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color.alpha_channel).to eq(0)
  end

  it 'gradient_text.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/gradient_text.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.type).to eq(:gradient)
  end

  it 'gradient_text.docx - outline check' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/gradient_text.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.type).to eq(:none)
  end

  # File with created with 'no fill' fill
  it 'text_with_out_fill_color.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/text_with_out_fill_color.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color).to eq(:none)
  end

  # Take different
  it 'color_fills_nothing.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/color_fills_nothing.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color).to eq(:none)
  end

  it 'without_outline_color.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/without_outline_color.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
  end

  it 'without_outline_width.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/without_outline_color.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to be_zero
  end

  it 'without_outline.docx - color outline check' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/text_art/without_outline_color.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
  end
end
