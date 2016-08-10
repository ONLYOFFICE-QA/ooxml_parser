require_relative 'chart/chart'
require_relative 'drawing/docx_drawing'
require_relative 'shape/shape'
require_relative 'picture/old_docx_picture'
# Class for storing fallback graphic elements
module OoxmlParser
  class AlternateContent < OOXMLDocumentObject
    attr_accessor :office2010_content, :office2007_content

    def self.parse(alt_content_node, parent: nil)
      alternate_content = AlternateContent.new
      alternate_content.parent = parent
      alt_content_node.xpath('*').each do |alternate_content_node_child|
        begin
          alternate_content_node_child.xpath('w:drawing')
        rescue Nokogiri::XML::XPath::SyntaxError # This mean it is Chart
          case alternate_content_node_child.name
          when 'Choice'
            alternate_content.office2010_content = Office2010ChartStyle.parse(alternate_content_node_child)
          when 'Fallback'
            alternate_content.office2007_content = Office2007ChartStyle.parse(alternate_content_node_child)
          end
          next
        end
        case alternate_content_node_child.name
        when 'Choice'
          alternate_content.office2010_content = DocxDrawing.parse(alternate_content_node_child.xpath('w:drawing').first, parent: alternate_content) unless alternate_content_node_child.xpath('w:drawing').first.nil?
        when 'Fallback'
          alternate_content.office2007_content = OldDocxPicture.parse(alternate_content_node_child.xpath('w:pict').first, parent: alternate_content) unless alternate_content_node_child.xpath('w:pict').first.nil?
        end
      end
      alternate_content
    end
  end
end
