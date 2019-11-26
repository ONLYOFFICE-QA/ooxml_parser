# frozen_string_literal: true

module OoxmlParser
  # Class with data of Note
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
      doc = note.parse_xml(note.file_path(params[:target]))
      if note.type.include?('footer')
        xpath_note = '//w:ftr'
      elsif note.type.include?('header')
        xpath_note = '//w:hdr'
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
          end
        end
      end
      note
    end

    # @param target [String] name of target
    # @return [String] path to note xml file
    def file_path(target)
      file = "#{OOXMLDocumentObject.path_to_folder}word/#{target}"
      return file if File.exist?(file)

      "#{OOXMLDocumentObject.path_to_folder}#{target}" unless File.exist?(file)
    end
  end
end
