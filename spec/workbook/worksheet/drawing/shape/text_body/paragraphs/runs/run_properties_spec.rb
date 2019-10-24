# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'fill_without_outline.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/fill_without_outline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color).to eq(OoxmlParser::Color.new(79, 129, 189))
  end

  it 'fill_without_outline.xlsx - type check' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/fill_without_outline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.type).to eq(:solid)
  end

  it 'fill_without_outline.xlsx - nil size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/fill_without_outline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to be_zero
  end

  it 'outline_without_fill.xlsx - color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/outline_without_fill.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.color_scheme.color).to eq(OoxmlParser::Color.new(149, 55, 53))
  end

  it 'outline_without_fill.xlsx - size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/outline_without_fill.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to eq(OoxmlParser::OoxmlSize.new(1.24, :point))
  end

  it 'outline_width_6px.xlsx - big size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/outline_width_6px.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to eq(OoxmlParser::OoxmlSize.new(6, :point))
  end

  it 'gradient_fill.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/gradient_fill.xlsx')
    font_color = xlsx.worksheets[0].drawings
                     .first.shape.text_body
                     .paragraphs.first.runs
                     .first.properties.font_color
    expect(font_color.type).to eq(:gradient)
    expect(font_color.color.gradient_stops
               .first.color.converted_color)
      .to eq(OoxmlParser::Color.new(179, 162, 199))
    expect(font_color.color.gradient_stops.first.position).to eq(0)
    expect(font_color.color.gradient_stops[1]
               .color.converted_color)
      .to eq(OoxmlParser::Color.new(250, 192, 144))
    expect(font_color.color.gradient_stops[1].position).to eq(100)
  end

  it 'linear_gradient_no_scaled' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/linear_gradient_no_scaled.xlsx')
    expect(xlsx.worksheets[1].drawings[34].shape.text_body.paragraphs.first.runs.first.properties.font_color.color.linear_gradient.scaled).to be_nil
  end

  it 'shape_with_link.xlsx Link To' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/shape_with_link.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.hyperlink).not_to be_nil
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.hyperlink.url).to eq('http://www.yandex.ru/')
  end

  it 'baseline_subscript' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/properties/baseline_subscript.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.baseline).to eq(:subscript)
  end
end
