# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  merged_cells = OoxmlParser::Parser.parse('spec/presentation/slide/graphic_frame' \
                                           '/graphic_data/table/rows/cells/' \
                                           'properties/merged_cells.pptx')

  it 'TableStyleBackgroundColor' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/table_style_background_color.pptx')
    table = pptx.slides.first.elements.last.graphic_data.first
    expect(table.rows.first.cells.first.properties.color).to be_nil
    expect(table.rows.first.cells[1].properties.color.color.converted_color).to eq(OoxmlParser::Color.new(68, 114, 196))
  end

  it 'table_border_style.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/table_border_style.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    table.rows.each do |current_row|
      current_row.cells.each do |current_cell|
        current_cell.properties.borders.each_side do |current_side|
          expect(current_side.fill.color.converted_color).to eq(OoxmlParser::Color.new(165, 165, 165))
          expect(current_side.width).to eq(OoxmlParser::OoxmlSize.new(4.5, :point))
        end
        expect(current_cell.properties.color.color.converted_color).to eq(OoxmlParser::Color.new(68, 114, 196))
      end
    end
  end

  it 'table_cell_vertical_align_top.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/table_cell_vertical_align_top.pptx')
    first_row = pptx.slides.first.elements.last.graphic_data.first.rows.first
    expect(first_row.cells.first.properties.anchor).to eq(:top)
  end

  it 'table_cell_vertical_align_bottom.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/table_cell_vertical_align_bottom.pptx')
    first_row = pptx.slides.first.elements.last.graphic_data.first.rows.first
    expect(first_row.cells.first.properties.anchor).to eq(:bottom)
  end

  it 'table_background_color.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/table_background_color.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    expect(OoxmlParser::Color.to_color(table.rows.first.cells.first.properties.color)).to eq(OoxmlParser::Color.new(70, 93, 73))
  end

  it 'text_direction' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/graphic_frame/graphic_data/table/rows/cells/properties/text_direction.pptx')
    expect(pptx.slides.first.elements.last.graphic_data.first.rows.first.cells.first.properties.text_direction).to eq(:horz)
  end

  describe 'Merge properties' do
    it 'grid span is correct' do
      expect(merged_cells.slides.first.graphic_frames
                 .first.graphic_data.first
                 .rows.first.cells.first.grid_span).to eq(2)
    end

    it 'horizontal_merge is correct' do
      expect(merged_cells.slides.first.graphic_frames
                 .first.graphic_data.first
                 .rows.first.cells[1].horizontal_merge).to eq(1)
    end

    it 'vertical_merge is correct' do
      expect(merged_cells.slides.first.graphic_frames
                 .first.graphic_data.first
                 .rows[2].cells[0].vertical_merge).to eq(1)
    end
  end
end
