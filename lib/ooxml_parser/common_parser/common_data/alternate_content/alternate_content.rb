require_relative 'alternate_content/choice'
require_relative 'alternate_content/chart_style'
require_relative 'chart/chart'
require_relative 'drawing/docx_drawing'
require_relative 'picture/old_docx_picture'
module OoxmlParser
  # Class for storing fallback graphic elements
  class AlternateContent < OOXMLDocumentObject
    attr_accessor :office2010_content, :office2007_content
    # @return [Choice] choice data
    attr_accessor :choice

    # Parse AlternateContent
    # @param [Nokogiri::XML:Node] node with Relationships
    # @return [AlternateContent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        begin
          node_child.xpath('w:drawing')
        rescue Nokogiri::XML::XPath::SyntaxError # This mean it is Chart
          case node_child.name
          when 'Choice'
            @office2010_content = ChartStyle.new(parent: self).parse(node_child)
          when 'Fallback'
            @office2007_content = ChartStyle.new(parent: self).parse(node_child)
          end
          next
        end
        case node_child.name
        when 'Choice'
          @office2010_content = DocxDrawing.new(parent: self).parse(node_child.xpath('w:drawing').first) unless node_child.xpath('w:drawing').first.nil?
          @choice = Choice.new(parent: self).parse(node_child)
        when 'Fallback'
          @office2007_content = OldDocxPicture.parse(node_child.xpath('w:pict').first, parent: self) unless node_child.xpath('w:pict').first.nil?
        end
      end
      self
    end
  end
end
