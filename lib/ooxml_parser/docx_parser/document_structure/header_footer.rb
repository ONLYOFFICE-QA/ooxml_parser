# frozen_string_literal: true

module OoxmlParser
  # Class Header Footer classes
  class HeaderFooter < OOXMLDocumentObject
    # @return [Integer] id of header-footer
    attr_reader :id
    # @return [Array<OOXMLDocumentObject>] list of elements if object
    attr_reader :elements
    # @return [Symbol] `:header` or `:footer`
    attr_reader :type
    # @return [String] suffix for object files
    attr_reader :path_suffix

    def initialize(type = :header, parent: nil)
      @type = type
      @elements = []
      super(parent: parent)
    end

    # @return [String] string for search of xpath
    def xpath_for_search
      "//w:#{path_suffix}"
    end

    # @return [String] string with xml path
    def xml_path
      "word/#{path_suffix}s.xml"
    end

    # Parse type and path suffix by node type
    # @param [Nokogiri::XML:Node] node with HeaderFooter
    def parse_type(node)
      case node.name
      when 'footnoteReference'
        @path_suffix = 'footnote'
        @type = :header
      when 'endnoteReference'
        @path_suffix = 'endnote'
        @type = :footer
      end
    end

    # Parse HeaderFooter
    # @param [Nokogiri::XML:Node] node with HeaderFooter
    # @return [HeaderFooter] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        end
      end
      parse_type(node)
      doc = parse_xml(root_object.unpacked_folder + xml_path)
      doc.search(xpath_for_search).each do |footnote|
        next unless footnote.attribute('id').value.to_i == @id

        paragraph_number = 0
        footnote.xpath('w:p').each do |paragraph|
          @elements << DocumentStructure.default_paragraph_style.dup.parse(paragraph, paragraph_number, DocumentStructure.default_run_style, parent: self)
          paragraph_number += 1
        end
      end
      self
    end
  end
end
