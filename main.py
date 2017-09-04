from peewee import *
import wechatsogou
import datetime
import gspreadsheet

db = SqliteDatabase('wechat.sqlite')

# Add official accounts your are following below
gzhs_wechat_id = [
    'ninechapter', 'collegedaily', 'iProgrammer'
]

class GZH(Model):
    # Fld types in peewee: http://docs.peewee-orm.com/en/latest/peewee/models.html#field-types-table
    open_id = TextField(null=True)
    profile_url = TextField()
    qrcode = TextField()
    post_perm = IntegerField()
    wechat_id = CharField()
    wechat_name = TextField()
    introduction = TextField()
    authentication = TextField()
    head_image = TextField()
    created_at = DateTimeField()

    class Meta:
        database = db


class Article(Model):
    title = TextField()
    abstract = TextField()
    send_id = DoubleField()
    type = CharField()
    source_url = TextField()
    cover = TextField()
    author = TextField()
    main = IntegerField()
    fileid = DoubleField()
    copyright_stat = IntegerField()
    created_at = DateTimeField()
    content_url = TextField()
    is_valid = BooleanField()
    updated_at = DateTimeField()
    gzh = ForeignKeyField(GZH, 'articles')

    class Meta:
        database = db  # This model uses the "people.db" database.


api = wechatsogou.WechatSogouAPI()  # https://github.com/Chyroc/WechatSogou
db.connect()
db.create_tables([GZH, Article], True)


def get_recent_articles(api, wechat_name):
    res = api.get_gzh_artilce_by_history(wechat_name)
    articles = res['article']
    return articles


def get_valid_articles(gzh_id=None):
    if gzh_id is None:
        return Article.select().where(Article.is_valid)
    valid_articles = Article.select().where(Article.gzh == gzh_id, Article.is_valid)
    if valid_articles.exists():
        return valid_articles
    else:
        return []


def get_titles_from_valid_articles_in_gzh(gzh_id):
    valid_article_titles = set()
    valid_articles = get_valid_articles(gzh_id)
    for valid_article in valid_articles:
        valid_article_titles.add(valid_article.title)
    return valid_article_titles


def create_gzh(doc_main, doc_arc, api, wechat_id):
    gzh_results = api.search_gzh(wechat_id)
    for gzh in gzh_results:
        if gzh['wechat_id'] == wechat_id:
            gzh = GZH.create(                         # create gzh in db
                authentication=gzh['authentication'],
                head_image=gzh['headimage'],
                introduction=gzh['introduction'],
                post_perm=gzh['post_perm'],
                profile_url=gzh['profile_url'],
                qrcode=gzh['qrcode'],
                wechat_id=gzh['wechat_id'],
                wechat_name=gzh['wechat_name'],
                created_at=datetime.datetime.now()
            )
            doc_main.create_new_sheet_for_new_gzh(gzh.wechat_name) # create sheet for gzh
            doc_arc.create_new_sheet_for_new_gzh(gzh.wechat_name)  # create sheet for gzh in doc_arc
            return gzh


def get_zsh_info(doc_main, doc_arc, api, gzh_wechat_id):
    query = GZH.select().where(GZH.wechat_id == gzh_wechat_id)
    if not query.exists():
        return create_gzh(doc_main, doc_arc, api, gzh_wechat_id)
    return query.get()


def insert_new_article(gzh_id, article):
    Article.create(
        abstract=article['abstract'],
        send_id=article['send_id'],
        type=article['type'],
        source_url=article['source_url'],
        cover=article['cover'],
        author=article['author'],
        main=article['main'],
        title=article['title'],
        fileid=article['fileid'],
        copyright_stat=article['copyright_stat'],
        created_at=datetime.datetime.fromtimestamp(article['datetime']),
        content_url=article['content_url'],
        is_valid=True,
        gzh=gzh_id,
        updated_at=datetime.datetime.now()
    )


def update_article_in_db(recent_article):
    query = Article.update(
        content_url=recent_article['content_url'],
        updated_at=datetime.datetime.now()
    ).where(Article.title == recent_article['title'])
    query.execute()


# remove articles not fetchable by the api
# 1. invalidate them in db
# 2. move those rows of articles to an archived spreadsheet
def expire_old_articles(doc_main, doc_arc, old_article_titles, gzh_id):
    gzh = GZH.get(GZH.id == gzh_id)
    sheet_cur = doc_main    .worksheets[gzh.wechat_name]
    sheet_arc = doc_arc.worksheets[gzh.wechat_name]
    for title in old_article_titles:
        invalidate_article_in_db_by_title(title, gzh_id)
        archive_expired_article_in_worksheet(sheet_cur, sheet_arc, title)


# 1. get the row of expired article (search with title regex) in worksheet
# 2. remove that row in worksheet
# 3. prepend that row to worksheet_arc
def archive_expired_article_in_worksheet(worksheet, worksheet_arc, title):
    # if not using re.compile('regex str') in find(), query string must 100% match text in cell
    row_num = worksheet.find(title)[0].row
    article_row = worksheet.get_row(row_num)
    worksheet.delete_rows(row_num, 1)
    worksheet_arc.insert_rows(1, 1, article_row)


def invalidate_article_in_db_by_title(title, gzh_id):
    article = Article.get(Article.title == title, Article.gzh == gzh_id)
    article.is_valid = False
    article.save()


def filter_duplicate_articles(articles):
    m = {}
    res = []
    for article in articles:
        title = article['title']
        if not m.get(title):
            m[title] = []
        m[title].append(article)

    for title, article_list in m.items():
        latest_article = article_list[0]
        if len(article_list) > 1:
            for article in article_list:
                if article['datetime'] > latest_article['datetime']:
                    latest_article = article
        res.append(latest_article)

    return res


def main():
    doc_main = gspreadsheet.GSpreadSheet('credentials.json', 'WeChat')
    doc_arc = gspreadsheet.GSpreadSheet('credentials.json', 'WeChat Archived')

    for gzh_wechat_id in gzhs_wechat_id:
        gzh = get_zsh_info(doc_main, doc_arc, api, gzh_wechat_id)  # 获取 gzh in db / 创建 in db & doc
        recent_articles = filter_duplicate_articles(get_recent_articles(api, gzh.wechat_name))

        valid_article_titles = get_titles_from_valid_articles_in_gzh(gzh.id)
        new_articles = []
        for recent_article in recent_articles:
            if recent_article['title'] not in valid_article_titles:
                new_articles.append(recent_article)
                insert_new_article(gzh.id, recent_article)
            else:
                update_article_in_db(recent_article)  # update article content_url and timestamp
                doc_main.update_article_url(gzh.wechat_name, recent_article)
                valid_article_titles.remove(recent_article['title'])

        if len(new_articles) != 0:  # insert new articles into sheet
            doc_main.add_new_articles(gzh.wechat_name, new_articles)

        # all left in valid_article_titles are not valid any more, since we removed all valid
        # ones (in recent_articles), others become history articles and can't be updated
        expire_old_articles(doc_main, doc_arc, valid_article_titles, gzh.id)

    print('done')

main()
