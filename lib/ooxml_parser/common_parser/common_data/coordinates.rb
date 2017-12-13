module OoxmlParser
  class Coordinates
    attr_accessor :row, :column, :list

    def initialize(row = nil, column = nil, list = nil)
      @row = row.nil? ? row : row.to_i
      @column = column
      @list = list
    end

    class << self
      # @param [String] string to parse
      # @return [Coordinates]
      def parse_coordinates_from_string(string)
        coordinates = Coordinates.new
        begin
          coordinates.list = string.match(/'\w+'/)[0].delete("''")
        rescue StandardError
          'Raise Exception'
        end
        string = string.split('!').last
        if coordinates?(string)
          coordinates.row = string.scan(/[0-9]/).join('').to_i
          coordinates.column = string.scan(/[A-Z]/).join('')
        end
        coordinates
      end

      def parser_coordinates_range(arguments_string)
        sheet_name = 'Sheet1'

        sheet_name, arguments_string = arguments_string.split('!') if arguments_string.include?('!')

        range = arguments_string.split(':')

        difference = []
        symbols_from = range.first.scan(/[a-zA-z]/).join
        symbols_to = range.last.scan(/[a-zA-z]/).join
        digits_from = range.first.scan(/[0-9]/).join
        digits_to = range.last.scan(/[0-9]/).join

        difference[0] = [symbols_from, symbols_to] unless symbols_from == symbols_to
        difference[1] = [digits_from, digits_to] unless digits_from == digits_to
        arguments_array = []
        case difference.length
        when 0
          arguments_array << Coordinates.new(digits_from, symbols_from)
        when 1
          (difference.first.first..difference.first.last).each do |symbol|
            arguments_array << Coordinates.new(digits_from, symbol, sheet_name)
          end
        when 2
          case difference.first
          when nil
            (difference.last.first..difference.last.last).each do |digit|
              arguments_array << Coordinates.new(digit, symbols_from, sheet_name)
            end
          else
            (difference.first.first..difference.first.last).each do |symbol|
              (difference.last.first..difference.last.last).each do |digit|
                arguments_array << Coordinates.new(digit, symbol, sheet_name)
              end
            end
          end
        else
          raise 'Wrong arguments format'
        end
        arguments_array
      end

      # This method check is argument contains coordinate
      # @param [String] string
      def coordinates?(string)
        !/^([A-Z]+)(\d+)$/.match(string).nil?
      end

      # @param [String or Fixnum] number to convert
      # @return [String] column's name
      # This method takes string like '12' or '45' etc.
      # and converts into spreadsheet column's name
      #    StringHelper.get_column_name('12')  #=> "L"
      #    StringHelper.get_column_name('45')  #=> "AS"
      #    StringHelper.get_column_name('287')  #=> "KA"
      def get_column_name(number)
        (number.to_i > 0 ? ('A'..'Z').to_a[(number.to_i - 1) % 26] + get_column_name((number.to_i - 1) / 26).reverse : '').reverse
      end
    end

    # get_column_number -> integer
    # @return [Integer] number of column
    # This method takes @column string
    # and converts into integer
    def get_column_number
      @column.reverse.each_char.reduce(0) do |result, char|
        result + (char.downcase.ord - 96) * (26**@column.reverse.index(char))
      end
    end

    # Compares rows of two cells
    # @param [Coordinates] other_cell other cell coordinates
    # @return [true, false] true, if row greater, than other row
    def row_greater_that_other?(other_cell)
      @row > other_cell.row
    end

    # Compares columns of two cells
    # @param [Coordinates] other_cell other cell coordinates
    # @return [true, false] true, if column greater, than other row
    def column_greater_that_other?(other_cell)
      get_column_number > other_cell.get_column_number
    end

    def to_s
      "#{@column}#{@row} #{@list ? "list: #{@list}" : ''}"
    end

    def nil?
      @column.nil? && @list.nil? && @row.nil?
    end

    def ==(other)
      other.is_a?(Coordinates) ? (@row == other.row && @column == other.column) : false
    end
  end
end
