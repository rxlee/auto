import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader, TensorDataset
from torchvision import transforms
import fiftyone as fo
import fiftyone.zoo as foz
import openai

available_datasets = foz.list_zoo_datasets()

print(available_datasets)
fo.config.default_ml_backend = "torch"


class MyModel(nn.Module):
    def __init__(self):
        super(MyModel, self).__init__()
        self.conv1 = nn.Conv2d(3, 32, kernel_size=3, stride=1, padding=1)
        self.conv2 = nn.Conv2d(32, 64, kernel_size=3, stride=1, padding=1)
        self.fc1 = nn.Linear(64 * 8 * 8, 128)
        self.fc2 = nn.Linear(128, 10)

    def forward(self, x):
        x = nn.functional.relu(self.conv1(x))
        x = nn.functional.max_pool2d(x, kernel_size=2, stride=2)
        x = nn.functional.relu(self.conv2(x))
        x = nn.functional.max_pool2d(x, kernel_size=2, stride=2)
        x = x.view(-1, 64 * 8 * 8)
        x = nn.functional.relu(self.fc1(x))
        x = self.fc2(x)
        return x


model = MyModel()

criterion = nn.CrossEntropyLoss()
optimizer = optim.Adam(model.parameters(), lr=0.001)

transform = transforms.Compose([transforms.ToTensor(),
                                transforms.Normalize((0.5, 0.5, 0.5), (0.5, 0.5, 0.5))])

data = torch.randn(100, 3, 32, 32)
labels = torch.randint(0, 10, (100,))

dataset = TensorDataset(data, labels)

num_samples = len(dataset)
split = int(num_samples * 0.8)
train_set, val_set = torch.utils.data.random_split(dataset, [split, num_samples - split])

train_loader = DataLoader(train_set, batch_size=32, shuffle=True)
val_loader = DataLoader(val_set, batch_size=32, shuffle=True)

response = openai.Image.create(
    prompt="a white siamese cat",
    n=1,
    size="1024x1024"
)
image_url = response['data'][0]['url']

response = openai.Image.create_edit(
    image=open("sunlit_lounge.png", "rb"),
    mask=open("mask.png", "rb"),
    prompt="A sunlit indoor lounge area with a pool containing a flamingo",
    n=1,
    size="1024x1024"
)
image_url = response['data'][0]['url']

response = openai.Image.create_variation(
    image=open("corgi_and_cat_paw.png", "rb"),
    n=1,
    size="1024x1024"
)
image_url = response['data'][0]['url']

import fiftyone as fo
import fiftyone.zoo as foz

print(foz.list_zoo_datasets())
dataset = foz.load_zoo_dataset("coco-2017", split="validation")
dataset.name = "coco-2017-validation-example"
dataset.persistent = True

session = fo.launch_app(dataset)

from io import BytesIO

byte_stream: BytesIO = BytesIO(10000)
byte_array = byte_stream.getvalue()
response = openai.Image.create_variation(
    image=byte_array,
    n=1,
    size="1024x1024"
)

from io import BytesIO
from PIL import Image

image = Image.open("image.png")
width, height = 256, 256
image = image.resize((width, height))

byte_stream = BytesIO()
image.save(byte_stream, format='PNG')
byte_array = byte_stream.getvalue()

response = openai.Image.create_variation(
    image=byte_array,
    n=1,
    size="1024x1024"
)
try:
    openai.Image.create_variation(
        open("image.png", "rb"),
        n=1,
        size="1024x1024"
    )
    print(response['data'][0]['url'])
except openai.error.OpenAIError as e:
    print(e.http_status)
print(e.error)

dataset = foz.download_zoo_dataset("cifar10", split="train")
