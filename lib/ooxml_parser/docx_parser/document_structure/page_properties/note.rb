# frozen_string_literal: true

module OoxmlParser
  # Class with data of Note
  class Note < OOXMLDocumentObject
    attr_accessor :type, :elements, :assigned_to

    def initialize(type: 'unknown')
      @elements = []
      @type = type
      super(parent: nil)
    end

    # @return [String] xpath for note in xml object
    def note_base_xpath
      @note_base_xpath ||= if @type.include?('footer')
                             '//w:ftr'
                           elsif @type.include?('header')
                             '//w:hdr'
                           else
                             raise NameError, "Unknown note type: #{@type}"
                           end
    end

    # Parse note data
    # @param params [Hash] data to parse
    # @return [Note] result of parsing
    def parse(params)
      @type = params[:type]
      @assigned_to = params[:assigned_to]
      @parent = params[:parent]
      doc = parse_xml(file_path(params[:target]))
      doc.search(note_base_xpath).each do |ftr|
        number = 0
        ftr.xpath('*').each do |sub_element|
          case sub_element.name
          when 'p'
            @elements << params[:default_paragraph].dup.parse(sub_element, number, params[:default_character], parent: self)
            number += 1
          when 'tbl'
            @elements << Table.new(parent: self).parse(sub_element, number)
            number += 1
          end
        end
      end
      self
    end

    # @param target [String] name of target
    # @return [String] path to note xml file
    def file_path(target)
      file = "#{root_object.unpacked_folder}word/#{target}"
      return file if File.exist?(file)

      "#{root_object.unpacked_folder}#{target}" unless File.exist?(file)
    end
  end
end
