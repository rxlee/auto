


# 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


# This is a sample Python script.
import openai
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import urllib.request

def download_img(img_url):
    request = urllib.request.Request(img_url)
    try:
        response = urllib.request.urlopen(request)
        img_name = "img.png"
        if (response.getcode() == 200):
            with open(img_name, "wb") as f:
                f.write(response.read()) # 将内容写入图片
            return img_name
    except:
        return "failed"

def print_hi():
    openai.api_key = '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'
    models = openai.Model.list()
    print(models.data[0].id)
    response = openai.Image.create(
        prompt="三只可爱小熊猫",
        n=1,
        size="500x500"
    )
    image_url = response['data'][0]['url']
    download_img(image_url)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
