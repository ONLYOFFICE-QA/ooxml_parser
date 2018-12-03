module OoxmlParser
  # Class Header Footer classes
  class HeaderFooter < OOXMLDocumentObject
    attr_accessor :id, :elements, :type
    attr_accessor :path_suffix

    def initialize(type = :header, parent: nil)
      @type = type
      @elements = []
      @parent = parent
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
      else
        warn "Unknown HeaderFooter type: #{node.name}"
      end
    end

    # Parse HeaderFooter
    # @param [Nokogiri::XML:Node] node with HeaderFooter
    # @return [HeaderFooter] result of parsing
    def parse(node)
      @id = node.attribute('id').value.to_i
      parse_type(node)
      doc = parse_xml(OOXMLDocumentObject.path_to_folder + xml_path)
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
