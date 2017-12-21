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
          @date = parse_date(value.value.to_s)
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

    private

    # Parse date and handle incorrect dates
    # @param value [Sting] value of date
    # @return [DateTime, String] if date correct or incorrect
    def parse_date(value)
      DateTime.parse(value)
    rescue ArgumentError
      warn "Date #{value} is incorrect format"
      value
    end
  end
end
