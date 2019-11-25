# frozen_string_literal: true

require 'spec_helper'

describe 'Slide' do
  it 'Exception on transform_of_object unknown' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/image_horizontal_alignment_center.pptx')
    expect { pptx.slides.first.content_horizontal_align(:slide, pptx.slide_size) }
      .to raise_error(RuntimeError, /Dont know this type object/)
  end

  describe 'Horizontal Align' do
    it 'image_horizontal_alignment_center' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/image_horizontal_alignment_center.pptx')
      expect(pptx.slides.first.content_horizontal_align(:image, pptx.slide_size)).to eq(:center)
    end

    it 'table_center_horizontal_align' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/table_right_horizontal_align.pptx')
      expect(pptx.slides.first.content_horizontal_align(:table, pptx.slide_size)).to eq(:right)
    end

    it 'content_horizontal_align unknwon' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape_unknown_location.pptx')
      expect(pptx.slides.first.content_horizontal_align(:shape, pptx.slide_size)).to eq(:unknown)
    end
  end

  describe 'Vertical Align' do
    it 'image_vertical_alignment_middle' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/image_vertical_alignment_middle.pptx')
      expect(pptx.slides.first.content_vertical_align(:image, pptx.slide_size)).to eq(:middle)
    end

    it 'shape_bottom_vertical_align' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape_bottom_vertical_align.pptx')
      expect(pptx.slides.first.content_vertical_align(:shape, pptx.slide_size)).to eq(:bottom)
    end

    it 'content_vertical_align unknwon' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape_unknown_location.pptx')
      expect(pptx.slides.first.content_vertical_align(:shape, pptx.slide_size)).to eq(:unknown)
    end
  end

  describe 'Content Distribute' do
    it 'chart_content_distribute_both' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/chart_content_distribute_vertically.pptx')
      expect(pptx.slides.first.content_distribute(:chart, pptx.slide_size)).to include(:vertically)
      expect(pptx.slides.first.content_distribute(:chart, pptx.slide_size)).to include(:horizontally)
    end

    it 'conent_distribue horizontally' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape_only_horizontal.pptx')
      expect(pptx.slides.first.content_distribute(:shape, pptx.slide_size)).to eq([:horizontally])
    end

    it 'conent_distribue vertically' do
      pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape_only_verticaly.pptx')
      expect(pptx.slides.first.content_distribute(:shape, pptx.slide_size)).to eq([:vertically])
    end
  end
end
