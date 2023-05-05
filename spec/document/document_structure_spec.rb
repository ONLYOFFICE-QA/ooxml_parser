# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocumentStructure do
  numbering = OoxmlParser::Parser.parse('spec/document/document_structure/numbering.docx')
  table_header_footer = OoxmlParser::Parser.parse('spec/document/document_structure/table_header_footer.docx')
  shape_header_footer = OoxmlParser::Parser.parse('spec/document/document_structure/shape_header_footer.docx')
  text_header_footer = OoxmlParser::Parser.parse('spec/document/document_structure/text_header_footer.docx')

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

    it 'element_by_description for text in footer' do
      expect(text_header_footer.element_by_description(location: :footer,
                                                       type: :simple)
                 .first.nonempty_runs
                 .last.text).to eq('footer_paragraph')
    end
  end

  describe 'DocumentStructure#note_by_description' do
    it 'Raise error on unknown note type' do
      expect do
        table_header_footer.note_by_description(:error)
      end.to raise_error(/There isn't this type of the note/)
    end
  end

  describe 'DocumentStructure#recognize_numbering' do
    it 'able to recognize basic numbering' do
      expect(numbering.recognize_numbering.first).to eq('decimal')
    end
  end

  describe 'DocumentStructure#outline' do
    it 'able to recognize basic outline' do
      expect(numbering.outline.first).to eq('decimal')
    end
  end
end
