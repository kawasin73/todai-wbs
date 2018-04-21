require 'axlsx'

class Resolver
  def initialize(file)
    @file = file

    @cur = 0
    @sp = [0]
    @counter = [0]
    @list = [nil]
    @cache = false
  end

  def run!(result)
    p = Axlsx::Package.new
    @sheet = p.workbook.add_worksheet(:name => "WBS")
    @sheet.add_row ["レベル1", "レベル2", "レベル3", "レベル4", "レベル5"]

    File.open(@file) do |f|
      f.each_line do |line|
        parse(line)
      end
    end
    flush
    p.use_shared_strings = true
    p.serialize(result)
  end

  private

  def parse(line)
    sp = line.scan(/\A\s+/).first.to_s.length
    if sp == @sp.last.to_i
      if @cache
        flush
      end
      @counter[@cur] += 1
      @list[@cur] = to_item(line)
      @cache = true
    elsif sp > @sp.last.to_i
      @cur += 1
      @counter.append(1)
      @sp.append(sp)
      @list.append(to_item(line))
      @cache = true
    else
      if @cache
        flush
      end

      @list = @list[0, @cur]
      @sp = @sp[0, @cur]
      @counter = @counter[0, @cur]
      @cur -= 1
      parse(line)
    end
  end

  def flush
    @sheet.add_row(@list)
    # reset
    @list = @list.map {|_| nil}
    @cache = false
  end

  def to_item(line)
    "#{title} #{line.gsub(/\A\s+-/, '')}"
  end

  def title
    @counter.join('.')
  end
end

r = Resolver.new('./source.md')

r.run!('./result.xlsx')
