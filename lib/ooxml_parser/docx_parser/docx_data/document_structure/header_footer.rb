module OoxmlParser
  # Class Header Footer classes
  class HeaderFooter < OOXMLDocumentObject
    attr_accessor :id, :elements, :type
    attr_accessor :path_suffix

    def initialize(type = :header)
      @type = type
      @elements = []
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
    def self.parse(node)
      header_footer = HeaderFooter.new
      header_footer.id = node.attribute('id').value.to_i
      header_footer.parse_type(node)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + header_footer.xml_path))
      doc.search(header_footer.xpath_for_search).each do |footnote|
        next unless footnote.attribute('id').value.to_i == header_footer.id
        paragraph_number = 0
        footnote.xpath('w:p').each do |paragraph|
          header_footer.elements << DocxParagraph.parse(paragraph, paragraph_number, DocumentStructure.default_table_paragraph_style, DocumentStructure.default_table_run_style)
          paragraph_number += 1
        end
      end
      header_footer
    end
  end
end
