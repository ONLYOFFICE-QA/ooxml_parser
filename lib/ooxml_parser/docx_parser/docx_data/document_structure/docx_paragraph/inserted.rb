require 'date'

module OoxmlParser
  # Class for parsing `w:ins` tag - Inserted Run Content
  class Inserted < OOXMLDocumentObject
    # @return [Integer] id of inserted
    attr_accessor :id
    # @return [String] author of insert
    attr_accessor :author
    # @return [Date] date of insert
    attr_accessor :date
    # @return [String] id of user
    attr_accessor :user_id
    # @return [ParagraphRun] inserted run
    attr_accessor :run

    # Parse Inserted object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Inserted] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        when 'author'
          @author = value.value.to_s
        when 'date'
          @date = DateTime.parse(value.value.to_s)
        when 'oouserid'
          @user_id = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'r'
          @run = ParagraphRun.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
