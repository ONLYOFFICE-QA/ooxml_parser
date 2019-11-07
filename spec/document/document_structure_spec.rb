# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocumentStructure do
  let(:table_header_footer) { OoxmlParser::Parser.parse('spec/document/document_structure/table_header_footer.docx') }
  let(:shape_header_footer) { OoxmlParser::Parser.parse('spec/document/document_structure/shape_header_footer.docx') }

  describe 'DocumentStructure#element_by_description' do
    it 'element_by_description for incorrect location' do
      expect do
        table_header_footer.element_by_description(location: :error,
                                                   type: :table)
      end.to raise_error(/Wrong global location/)
    end

    it 'element_by_description for incorrect canvas type' do
      expect do
        table_header_footer.element_by_description(location: :canvas,
                                                   type: :error)
      end.to raise_error(/Wrong location/)
    end

    it 'element_by_description for incorrect header type' do
      expect do
        table_header_footer.element_by_description(location: :header,
                                                   type: :error)
      end.to raise_error(/Wrong location/)
    end

    it 'element_by_description for incorrect footer type' do
      expect do
        table_header_footer.element_by_description(location: :footer,
                                                   type: :error)
      end.to raise_error(/Wrong location/)
    end

    it 'element_by_description for table in header' do
      expect(table_header_footer.element_by_description(location: :header,
                                                        type: :table)
                 .first.nonempty_runs
                 .first.text).to eq('table_header')
    end

    it 'element_by_description for table in footer' do
      expect(table_header_footer.element_by_description(location: :footer,
                                                        type: :table)
                 .first.nonempty_runs
                 .last.text).to eq('table_footer')
    end

    it 'element_by_description for shape in header' do
      expect(shape_header_footer.element_by_description(location: :header,
                                                        type: :shape)
                 .first.nonempty_runs
                 .first.text).to eq('shape_header')
    end

    it 'element_by_description for shape in footer' do
      expect(shape_header_footer.element_by_description(location: :footer,
                                                        type: :shape)
                 .first.nonempty_runs
                 .last.text).to eq('shape_footer')
    end
  end
end
