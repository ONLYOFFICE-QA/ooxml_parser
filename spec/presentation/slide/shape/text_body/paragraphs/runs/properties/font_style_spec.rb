require 'spec_helper'

describe 'My behaviour' do
  it 'Check not underlined' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/font_style_not_underlined.pptx')
    font_style = pptx.slides[0].elements.last.text_body.paragraphs.first.characters.first.properties.font_style
    expect(font_style.underlined).to eq(:none)
  end

  it 'font_style_underlined_none.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/font_style_underlined_none.pptx')
    expect(pptx.slides[0].elements.last.text_body.paragraphs.first.characters.first.properties.font_style).to eq(OoxmlParser::FontStyle.new)
  end

  it 'strikeout.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/strikeout.pptx')
    expect(pptx.slides.first.elements.first.text_body.paragraphs.first.runs.first.properties.font_style.strike).to eq(:single)
  end

  it 'several_paragraphs_font_style.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/several_paragraphs_font_style.pptx')
    expect(pptx.slides.first.elements.first.text_body.paragraphs[0].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(false, false, OoxmlParser::Underline.new(:single), :none))
    expect(pptx.slides.first.elements.first.text_body.paragraphs[1].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(true))
    expect(pptx.slides.first.elements.first.text_body.paragraphs[2].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(false, false, OoxmlParser::Underline.new(:single), :none))
    expect(pptx.slides.first.elements.first.text_body.paragraphs[3].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(true))
  end

  it 'copy_style_in_table_with_bold.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/copy_style_in_table_with_bold.pptx')
    expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
  end

  it 'copy_style_in_table_without_bold.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/copy_style_in_table_without_bold.pptx')
    expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_falsey
  end

  it 'copy_style_in_table_with_bold_2_cell_with_bold_text_too.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/copy_style_in_table_with_bold_2_cell_with_bold_text_too.pptx')
    expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
    expect(pptx.slides[0].elements.last.graphic_data.first.rows[1].cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
  end

  it 'chart_title_bold.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/chart_title_bold.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.title.elements.first.runs.first.properties.font_style.bold).to eq(true)
  end

  it 'chart_title_without_bold.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_style/chart_title_without_bold.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.title.elements.first.runs.first.properties.font_style.bold).to eq(false)
  end
end
