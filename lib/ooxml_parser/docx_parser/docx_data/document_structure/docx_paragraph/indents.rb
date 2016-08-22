module OoxmlParser
  class Indents < OOXMLDocumentObject
    attr_accessor :first_line_indent, :left_indent, :right_indent, :hanging_indent

    def initialize(first_line_indent = 0.0, left_indent = 0.0, right_indent = 0.0, hanging_indent = nil)
      @first_line_indent = first_line_indent
      @left_indent = left_indent
      @right_indent = right_indent
      @hanging_indent = hanging_indent
      @units = :dxa
    end

    def ==(other)
      if self.class != NilClass && other.class != NilClass
        if @first_line_indent.to_f == other.first_line_indent.to_f &&
           @left_indent.to_f.round(2) == other.left_indent.to_f.round(2) &&
           @right_indent.to_f.round(1) == other.right_indent.to_f.round(1)
          return true if @hanging_indent.nil? || other.hanging_indent.nil?
          @hanging_indent.to_f == other.hanging_indent.to_f
        else
          false
        end
      else
        is_a?(NilClass) && other.is_a?(NilClass)
      end
    end

    def equal_with_round(other, delta = 0.02)
      if is_a?(NilClass) && other.is_a?(NilClass)
        true
      else
        @first_line_indent.to_f.round(2) == other.first_line_indent.to_f.round(2) &&
          @left_indent.to_f.round(2) == other.left_indent.to_f.round(2) &&
          @right_indent.to_f.round(2) == other.right_indent.to_f.round(2) &&
          (@hanging_indent.to_f.round(2) - other.hanging_indent.to_f.round(2)).to_f.abs < delta
      end
    end

    def to_s
      "first line indent: #{@first_line_indent}, left indent: #{@left_indent},
     right indent: #{@right_indent}, hanging indent: #{@hanging_indent}"
    end

    # Parse hash with indents
    # @return [Indents] parsed class
    def self.parse_indents_hash(hash)
      Indents.new(hash['first_line_indent'], hash['left_indent'], hash['right_indent'], hash['hanging_indent'])
    end

    # Increases indent, depending on clicks to "increase indent" button in canvas
    # @param [Integer] count count of clicks (default = 1)
    def increase_indent_by_button(count = 1)
      count.times { @left_indent += 1.25 }
    end

    # @return [Indents
    def copy
      Indents.new(@first_line_indent, @left_indent, @right_indent, @hanging_indent)
    end

    def round(digits_after_dot = 0)
      [:first_line_indent, :left_indent, :right_indent, :hanging_indent].each do |ind|
        send("#{ind}=", send(ind).round(digits_after_dot)) unless send(ind).nil?
      end
      self
    end

    # Parse Indents
    # @param [Nokogiri::XML:Element] node with Indents
    # @return [Indents] value of Indents
    def self.parse(node)
      indents = Indents.new
      unless node.attribute('firstLine').nil?
        indents.first_line_indent = (node.attribute('firstLine').value.to_f / 566.9).round(2)
      end
      unless node.attribute('left').nil?
        indents.left_indent = (node.attribute('left').value.to_f / 566.9).round(2)
      end
      unless node.attribute('right').nil?
        indents.right_indent = (node.attribute('right').value.to_f / 566.9).round(2)
      end
      unless node.attribute('hanging').nil?
        indents.hanging_indent = (node.attribute('hanging').value.to_f / 566.9).round(2)
      end
      indents
    end
  end
end
