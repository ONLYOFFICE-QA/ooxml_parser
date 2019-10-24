# frozen_string_literal: true

module OoxmlParser
  class Note < OOXMLDocumentObject
    attr_accessor :type, :elements, :assigned_to

    def initialize
      @elements = []
    end

    def self.parse(params)
      note = Note.new
      note.type = params[:type]
      note.assigned_to = params[:assigned_to]
      note.parent = params[:parent]
      doc = note.parse_xml(OOXMLDocumentObject.path_to_folder + "word/#{params[:target]}")
      if note.type.include?('footer')
        xpath_note = '//w:ftr'
      elsif note.type.include?('header')
        xpath_note = '//w:hdr'
      else
        raise "Cannot parse unknown Note type: #{type}"
      end
      doc.search(xpath_note).each do |ftr|
        number = 0
        ftr.xpath('*').each do |sub_element|
          if sub_element.name == 'p'
            note.elements << params[:default_paragraph].dup.parse(sub_element, number, params[:default_character], parent: note)
            number += 1
          elsif sub_element.name == 'tbl'
            note.elements << Table.new(parent: note).parse(sub_element, number)
            number += 1
          elsif sub_element.name == 'std'
            sub_element.xpath('w:p').each do |p|
              note.elements << params[:default_paragraph].copy.parse(p, number, params[:default_character])
              number += 1
            end
            sub_element.xpath('w:sdtContent').each do |sdt_content|
              sdt_content.xpath('w:p').each do |p|
                note.elements << params[:default_paragraph].copy.parse(p, number, params[:default_character])
                number += 1
              end
            end
          end
        end
      end
      note
    end
  end
end
